<#
.SYNOPSIS
  Imports people and assets from CSV files into the IT Asset Tracker lists.

.DESCRIPTION
  Start from the templates in the release's "templates" folder (Assets-template.csv, People-template.csv).
  Only "Name" is required; leave any other column blank or delete it.

  People CSV columns:  Name, JobTitle, Login, Location, Status (Employed / Terminated / Leave / Unknown)
  Assets CSV columns:  Name, Category, AssetType, Location, AssignedTo, SharedUse, Make, Model, SerialNumber,
                       IPAddress, Quantity, Status, InServiceYear, PurchaseDate, WarrantyEnd, ClassKey, Notes
    Category    Computers, Servers, Network, Printers, Telephony, Connectivity, Other
    Status      In Use, Spare, In Repair, Review, Return to Vendor, Retired
    AssignedTo  a person's Name or Login (people not found are created)
    Location    a location name (locations not found are created)
    ClassKey    a replacement class key from the app's Settings (laptop, desktop, ap, ...); blank = set later
    Dates       yyyy-mm-dd

  Duplicates: a row whose SerialNumber (or Name, when there's no serial) already exists is skipped,
  or updated with -Update. People are matched by Login, then Name. Run with -WhatIf first to preview.

.EXAMPLE
  ./Import-Assets.ps1 -SiteUrl https://contoso.sharepoint.com/sites/it -ClientId <client-id> -PeopleCsv .\People.csv -AssetsCsv .\Assets.csv -WhatIf
.EXAMPLE
  ./Import-Assets.ps1 -SiteUrl https://contoso.sharepoint.com/sites/it -ClientId <client-id> -AssetsCsv .\Assets.csv -Update
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [Parameter(Mandatory = $true)][string]$SiteUrl,
  [Parameter(Mandatory = $true)][string]$ClientId,
  [string]$AssetsCsv,
  [string]$PeopleCsv,
  [string]$Prefix = "IT ",
  [switch]$Update
)
$ErrorActionPreference = "Stop"
if (-not $AssetsCsv -and -not $PeopleCsv) { throw "Give -AssetsCsv and/or -PeopleCsv." }
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
$LLoc = "${Prefix}Locations"; $LPeople = "${Prefix}People"; $LAssets = "${Prefix}Assets"; $LClasses = "${Prefix}Replacement Classes"

function V($x) { if ($null -eq $x) { return $null }; $s = "$x".Trim(); if ($s -eq "") { $null } else { $s } }
function Key($x) { ("$x").Trim().ToLower() }
$stats = [ordered]@{ LocationsAdded = 0; PeopleAdded = 0; PeopleUpdated = 0; AssetsAdded = 0; AssetsUpdated = 0; AssetsSkipped = 0 }

$locId = @{}; foreach ($i in Get-PnPListItem -List $LLoc -PageSize 500) { $locId[(Key $i["Title"])] = $i.Id }
function Get-LocId($name) {
  $n = V $name; if (-not $n) { return $null }
  $k = Key $n; if ($locId.ContainsKey($k)) { return $locId[$k] }
  if ($PSCmdlet.ShouldProcess($n, "Add location")) { $locId[$k] = (Add-PnPListItem -List $LLoc -Values @{ Title = $n }).Id } else { $locId[$k] = -1 }
  $stats.LocationsAdded++; return $locId[$k]
}

$byLogin = @{}; $byName = @{}; $maxKey = 0
foreach ($i in Get-PnPListItem -List $LPeople -PageSize 500) {
  if (V $i["Login"]) { $byLogin[(Key $i["Login"])] = $i.Id }
  $byName[(Key $i["Title"])] = $i.Id
  if ("$($i['PersonKey'])" -match '^p-(\d+)$') { $maxKey = [Math]::Max($maxKey, [int]$Matches[1]) }
}
function Find-Person($s) { $k = Key $s; if ($byLogin.ContainsKey($k)) { return $byLogin[$k] }; if ($byName.ContainsKey($k)) { return $byName[$k] }; return $null }
function Add-Person($name, $vals) {
  $script:maxKey++; $vals.Title = $name; $vals.PersonKey = "p-$($script:maxKey)"
  if (-not $vals.ContainsKey("EmploymentStatus")) { $vals.EmploymentStatus = "Employed" }
  $id = -1; if ($PSCmdlet.ShouldProcess($name, "Add person")) { $id = (Add-PnPListItem -List $LPeople -Values $vals).Id }
  $byName[(Key $name)] = $id; if ($vals.Login) { $byLogin[(Key $vals.Login)] = $id }
  $stats.PeopleAdded++; return $id
}

# ---------------------------------------------------------------- people
if ($PeopleCsv) {
  foreach ($r in Import-Csv $PeopleCsv -Encoding UTF8) {
    $name = V $r.Name; if (-not $name) { continue }
    $v = @{}
    if (V $r.JobTitle) { $v.JobTitle = V $r.JobTitle }
    if (V $r.Login) { $v.Login = V $r.Login }
    if (V $r.Status) { $v.EmploymentStatus = V $r.Status }
    if (V $r.Location) { $v.PersonLocation = Get-LocId $r.Location }
    $id = $null; if ($v.Login) { $id = Find-Person $v.Login }; if (-not $id) { $id = Find-Person $name }
    if ($id) {
      if ($Update -and $v.Count -and $PSCmdlet.ShouldProcess($name, "Update person")) { Set-PnPListItem -List $LPeople -Identity $id -Values $v | Out-Null; $stats.PeopleUpdated++ }
    } else { Add-Person $name $v | Out-Null }
  }
}

# ---------------------------------------------------------------- assets
if ($AssetsCsv) {
  $classId = @{}; foreach ($i in Get-PnPListItem -List $LClasses -PageSize 500) { $classId[(Key $i["ClassKey"])] = $i.Id }
  $bySerial = @{}; $byAsset = @{}; $maxA = 0
  foreach ($i in Get-PnPListItem -List $LAssets -PageSize 500) {
    if (V $i["SerialNumber"]) { $bySerial[(Key $i["SerialNumber"])] = $i.Id }
    $byAsset[(Key $i["Title"])] = $i.Id
    if ("$($i['AssetKey'])" -match '^a(\d+)$') { $maxA = [Math]::Max($maxA, [int]$Matches[1]) }
  }
  $cats = "Computers", "Servers", "Network", "Printers", "Telephony", "Connectivity", "Other"
  $statuses = "In Use", "Spare", "In Repair", "Review", "Return to Vendor", "Retired"
  foreach ($r in Import-Csv $AssetsCsv -Encoding UTF8) {
    $name = V $r.Name; if (-not $name) { continue }
    $v = @{ Title = $name }
    foreach ($k in "AssetType", "SharedUse", "Make", "Model", "SerialNumber", "IPAddress", "Notes") { if (V $r.$k) { $v[$k] = V $r.$k } }
    if (V $r.Category) { $c = $cats | Where-Object { $_ -eq (V $r.Category) }; $v.Category = $(if ($c) { $c } else { "Other" }) }
    if (V $r.Status) { $s = $statuses | Where-Object { $_ -eq (V $r.Status) }; if ($s) { $v.AssetStatus = $s } else { Write-Warning "$name - unknown status '$($r.Status)', left as default" } }
    if (V $r.Quantity) { $v.Quantity = [int]$r.Quantity }
    if (V $r.InServiceYear) { $v.InServiceYear = [int]$r.InServiceYear; $v.YearSource = "Imported" }
    foreach ($k in "PurchaseDate", "WarrantyEnd") { if (V $r.$k) { $v[$k] = ([datetime](V $r.$k)).ToString("yyyy-MM-dd") } }
    if (V $r.Location) { $v.AssetLocation = Get-LocId $r.Location }
    if (V $r.AssignedTo) { $p = Find-Person $r.AssignedTo; if (-not $p) { $p = Add-Person (V $r.AssignedTo) @{} }; $v.AssignedPerson = $p }
    if (V $r.ClassKey) { if ($classId.ContainsKey((Key $r.ClassKey))) { $v.ReplacementClass = $classId[(Key $r.ClassKey)] } else { Write-Warning "$name - no replacement class '$($r.ClassKey)'" } }
    foreach ($k in @($v.Keys)) { if ($v[$k] -eq -1) { $v.Remove($k) } }   # -WhatIf placeholders

    $id = $null
    if ($v.SerialNumber -and $bySerial.ContainsKey((Key $v.SerialNumber))) { $id = $bySerial[(Key $v.SerialNumber)] }
    elseif (-not $v.SerialNumber -and $byAsset.ContainsKey((Key $name))) { $id = $byAsset[(Key $name)] }
    if ($id) {
      if ($Update) { if ($PSCmdlet.ShouldProcess($name, "Update asset")) { Set-PnPListItem -List $LAssets -Identity $id -Values $v | Out-Null }; $stats.AssetsUpdated++ }
      else { $stats.AssetsSkipped++ }
      continue
    }
    $maxA++; $v.AssetKey = "a{0:d4}" -f $maxA
    if (-not $v.ContainsKey("AssetStatus")) { $v.AssetStatus = "In Use" }
    if (-not $v.ContainsKey("Category")) { $v.Category = "Other" }
    if ($PSCmdlet.ShouldProcess($name, "Add asset")) { $newId = (Add-PnPListItem -List $LAssets -Values $v).Id } else { $newId = -1 }
    if ($v.SerialNumber) { $bySerial[(Key $v.SerialNumber)] = $newId }; $byAsset[(Key $name)] = $newId
    $stats.AssetsAdded++
  }
}

Write-Host ""
$stats.GetEnumerator() | ForEach-Object { "{0,-16} {1}" -f $_.Key, $_.Value } | Write-Host
if ($WhatIfPreference) { Write-Host "Preview only (-WhatIf). Nothing was changed." -ForegroundColor Yellow }
