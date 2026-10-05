<#
    AddToGroups.ps1
    Adds the computer the script runs on to one or more AD groups.

    TS step "Run Command Line" (Run this step as: delegated account):
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File AddToGroups.ps1 "group1":"group2":"group3"

    Exit codes:
        0 = OK (added or already member)
        2 = computer object not found
        3 = at least one group not found
        4 = at least one Add/IsMember failed (e.g. access denied)
        5 = setup / LDAP bind failed

    Log: %windir%\Temp\AddToGroups.log (plus stdout -> smsts.log)
#>
param(
    [Parameter(Mandatory, Position = 0)][string]$GroupList,
    [string]$LogPath = "$env:windir\Temp\AddToGroups.log"
)

function Log([string]$Msg) {
    $line = "{0:yyyy-MM-dd HH:mm:ss}  {1}" -f (Get-Date), $Msg
    Write-Output $line
    try { Add-Content -Path $LogPath -Value $line -ErrorAction Stop } catch {}
}

# ---------- Setup: explicit domain, DC and search root (no ADSI auto-detection) ----------
try {
    $Groups   = $GroupList.Split(':') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $Domain   = (Get-CimInstance Win32_ComputerSystem).Domain          # e.g. svo.energie
    $DomainDn = 'DC=' + ($Domain -replace '\.', ',DC=')                # e.g. DC=svo,DC=energie

    # DC holding the computer's secure channel (= where the join landed); fallback: domain FQDN
    $DC = $null
    $sc = nltest /sc_query:$Domain 2>$null | Select-String 'Trusted DC Name'
    if ($sc) { $DC = ($sc.ToString() -replace '.*\\', '').Trim() }
    if (-not $DC) { $DC = $Domain }

    $Root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DC/$DomainDn")
    $Root.RefreshCache()   # force bind now so LDAP/DNS problems show up here

    Log "Run as: $env:USERDOMAIN\$env:USERNAME | Computer: $env:COMPUTERNAME | DC: $DC | Root: $DomainDn"
    Log "Groups: $($Groups -join ', ')"
}
catch {
    Log "Setup/bind failed (DC '$DC', root '$DomainDn'): $($_.Exception.Message)"
    exit 5
}

function Find-Path([string]$Filter) {
    $s = New-Object System.DirectoryServices.DirectorySearcher($Root, $Filter)
    $s.SearchScope = 'Subtree'
    $r = $s.FindOne()
    if ($r) { $r.Path } else { $null }
}

# ---------- Computer object (retry up to ~60 s for replication) ----------
$ComputerPath = $null
for ($i = 1; $i -le 6 -and -not $ComputerPath; $i++) {
    try   { $ComputerPath = Find-Path "(&(objectCategory=computer)(sAMAccountName=$env:COMPUTERNAME`$))" }
    catch { Log "Computer lookup attempt ${i}: $($_.Exception.Message)" }
    if (-not $ComputerPath -and $i -lt 6) { Start-Sleep -Seconds 10 }
}
if (-not $ComputerPath) {
    Log "Computer '$env:COMPUTERNAME' not found on $DC"
    exit 2
}
Log "Computer: $ComputerPath"

# ---------- Groups ----------
$rc = 0
foreach ($Name in $Groups) {
    try {
        $GroupPath = Find-Path "(&(objectCategory=group)(sAMAccountName=$Name))"
        if (-not $GroupPath) {
            Log "Group '$Name' not found"
            $rc = [Math]::Max($rc, 3)
            continue
        }

        $Grp = [ADSI]$GroupPath
        if ($Grp.IsMember($ComputerPath)) {
            Log "Already member of '$Name'"
        }
        else {
            $Grp.Add($ComputerPath)
            Log "Added to '$Name'"
        }
    }
    catch {
        Log "Group '$Name': $($_.Exception.Message)"
        $rc = [Math]::Max($rc, 4)
    }
}

Log "Finished with exit code $rc"
exit $rc
