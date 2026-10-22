<#
.SYNOPSIS
  Creates (or upgrades) the SharePoint lists the IT Asset Tracker web part uses.

.DESCRIPTION
  Creates these lists on the site, linked to each other:
    IT Locations             sites / buildings
    IT Replacement Classes   device classes with unit cost and lifecycle (drives the budget)
    IT People                staff, with location and employment status
    IT Assets                every device, server, switch, printer...
    IT Actions               IT to-do list
    IT Asset Log             dated notes and assignment history per device
    IT Contracts             subscriptions, support renewals, circuits, leases

  Safe to run again: lists that already exist are kept, and any columns they're missing are added.
  Nothing is ever deleted.

  The replacement classes start with example costs and lifecycles - change them in the app's Settings.
  Use -Locations to create your sites up front, and -SampleData to load a fictional demo company
  ("Contoso Manufacturing") so you can try every feature before entering real data.

.REQUIREMENTS
  - PowerShell 7.4 or later (pwsh)
  - PnP.PowerShell:   Install-Module PnP.PowerShell -Scope CurrentUser
  - An Entra ID app registration for PnP interactive sign-in (one time per tenant):
        Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP PowerShell" -Tenant <yourtenant>.onmicrosoft.com
    Use the Client ID it prints with -ClientId.
  - Site Owner rights on the target site.
  - If Windows blocks the script after download:  Unblock-File .\Install-AssetTracker.ps1

.EXAMPLE
  ./Install-AssetTracker.ps1 -SiteUrl https://contoso.sharepoint.com/sites/it -ClientId <client-id>
.EXAMPLE
  ./Install-AssetTracker.ps1 -SiteUrl https://contoso.sharepoint.com/sites/it -ClientId <client-id> -Locations "Headquarters","Plant 2","Warehouse"
.EXAMPLE
  # A test site with demo data
  ./Install-AssetTracker.ps1 -SiteUrl https://contoso.sharepoint.com/sites/it-demo -ClientId <client-id> -SampleData
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$SiteUrl,
  [Parameter(Mandatory = $true)][string]$ClientId,
  [string]$Prefix = "IT ",
  [string[]]$Locations = @(),
  [switch]$SampleData,
  [string]$SampleFolder = (Join-Path (Split-Path $PSScriptRoot -Parent) "sample-data")
)

$ErrorActionPreference = "Stop"
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

$L = [ordered]@{
  Locations = "${Prefix}Locations"; Classes = "${Prefix}Replacement Classes"; People = "${Prefix}People"; Assets = "${Prefix}Assets"
  Actions = "${Prefix}Actions"; Log = "${Prefix}Asset Log"; Contracts = "${Prefix}Contracts"
}

# ---------------------------------------------------------------- helpers
function Ensure-List($title, $titleDisplay) {
  $l = Get-PnPList -Identity $title -ErrorAction SilentlyContinue
  if ($l) { Write-Host "List exists: $title" -ForegroundColor DarkGray; return $false }
  Write-Host "Creating list: $title" -ForegroundColor Cyan
  New-PnPList -Title $title -Template GenericList -EnableVersioning | Out-Null
  Set-PnPField -List $title -Identity "Title" -Values @{ Title = $titleDisplay } | Out-Null
  return $true
}
$script:have = @{}
function Has-Field($list, $internal) {
  if (-not $script:have.ContainsKey($list)) { $script:have[$list] = @{}; foreach ($f in Get-PnPField -List $list) { $script:have[$list][$f.InternalName] = $true } }
  return $script:have[$list].ContainsKey($internal)
}
function Add-FieldXml($list, $internal, $xml) {
  if (Has-Field $list $internal) { return }
  Add-PnPFieldFromXml -List $list -FieldXml $xml | Out-Null
  $script:have[$list][$internal] = $true
  Write-Host "  + $internal"
}
function F-Text($list, $n, $d, [switch]$Index) { Add-FieldXml $list $n "<Field Type='Text' DisplayName='$d' Name='$n' StaticName='$n' Indexed='$(if ($Index) {'TRUE'} else {'FALSE'})' />" }
function F-Note($list, $n, $d) { Add-FieldXml $list $n "<Field Type='Note' DisplayName='$d' Name='$n' StaticName='$n' NumLines='6' RichText='FALSE' />" }
function F-Num($list, $n, $d) { Add-FieldXml $list $n "<Field Type='Number' DisplayName='$d' Name='$n' StaticName='$n' Decimals='0' />" }
function F-Money($list, $n, $d) { Add-FieldXml $list $n "<Field Type='Currency' DisplayName='$d' Name='$n' StaticName='$n' Decimals='2' />" }
function F-Bool($list, $n, $d) { Add-FieldXml $list $n "<Field Type='Boolean' DisplayName='$d' Name='$n' StaticName='$n'><Default>0</Default></Field>" }
function F-Date($list, $n, $d, [switch]$Index) { Add-FieldXml $list $n "<Field Type='DateTime' Format='DateOnly' DisplayName='$d' Name='$n' StaticName='$n' Indexed='$(if ($Index) {'TRUE'} else {'FALSE'})' />" }
function F-Choice($list, $n, $d, [string[]]$choices, $default) {
  $c = ($choices | ForEach-Object { "<CHOICE>$_</CHOICE>" }) -join ""
  Add-FieldXml $list $n "<Field Type='Choice' DisplayName='$d' Name='$n' StaticName='$n' Format='Dropdown'><Default>$default</Default><CHOICES>$c</CHOICES></Field>"
}
function F-Lookup($list, $n, $d, $target, [switch]$Index) {
  $t = Get-PnPList -Identity $target
  Add-FieldXml $list $n "<Field Type='Lookup' DisplayName='$d' Name='$n' StaticName='$n' List='{$($t.Id)}' ShowField='Title' Indexed='$(if ($Index) {'TRUE'} else {'FALSE'})' />"
}
function Set-View($list, [string[]]$fields) {
  $v = Get-PnPView -List $list | Where-Object { $_.DefaultView } | Select-Object -First 1
  if ($v) { Set-PnPView -List $list -Identity $v.Id -Fields $fields | Out-Null }
}

# ---------------------------------------------------------------- lists and columns
$new = @{}
$new.Locations = Ensure-List $L.Locations "Location"

$new.Classes = Ensure-List $L.Classes "Class"
F-Text  $L.Classes "ClassKey" "Class Key" -Index
F-Text  $L.Classes "StandardModel" "Standard Replacement"
F-Money $L.Classes "UnitCost" "Unit Cost"
F-Num   $L.Classes "LifecycleYears" "Lifecycle (Years)"
F-Text  $L.Classes "CostSource" "Cost Source"
F-Num   $L.Classes "SpareMin" "Keep On Hand per Site"
if ($new.Classes) { Set-View $L.Classes @("LinkTitle", "ClassKey", "StandardModel", "UnitCost", "LifecycleYears", "SpareMin") }

$new.People = Ensure-List $L.People "Full Name"
F-Text   $L.People "PersonKey" "Person Key" -Index
F-Text   $L.People "JobTitle" "Job Title"
F-Text   $L.People "Login" "Login" -Index
F-Lookup $L.People "PersonLocation" "Location" $L.Locations -Index
F-Choice $L.People "EmploymentStatus" "Employment Status" @("Employed", "Terminated", "Leave", "Unknown") "Employed"
F-Bool   $L.People "O365" "O365"
F-Bool   $L.People "MFA" "MFA"
F-Bool   $L.People "VPN" "VPN"
F-Bool   $L.People "AV" "AV"
F-Text   $L.People "AlsoKnownAs" "Also Known As"
if ($new.People) { Set-View $L.People @("LinkTitle", "JobTitle", "Login", "PersonLocation", "EmploymentStatus") }

$new.Assets = Ensure-List $L.Assets "Asset Name"
F-Text   $L.Assets "AssetKey" "Asset Key" -Index
F-Choice $L.Assets "Category" "Category" @("Computers", "Servers", "Network", "Printers", "Telephony", "Connectivity", "Other") "Computers"
F-Text   $L.Assets "AssetType" "Type"
F-Lookup $L.Assets "AssetLocation" "Location" $L.Locations -Index
F-Lookup $L.Assets "AssignedPerson" "Assigned To" $L.People -Index
F-Text   $L.Assets "SharedUse" "Shared Use / Role"
F-Text   $L.Assets "Make" "Make"
F-Text   $L.Assets "Model" "Model"
F-Text   $L.Assets "SerialNumber" "Serial / Service Tag" -Index
F-Text   $L.Assets "IPAddress" "IP Address"
F-Num    $L.Assets "Quantity" "Quantity"
F-Choice $L.Assets "AssetStatus" "Status" @("In Use", "Spare", "In Repair", "Review", "Return to Vendor", "Retired") "In Use"
F-Text   $L.Assets "TransitionPlan" "Transition Plan"
F-Text   $L.Assets "Setup" "Setup"
F-Num    $L.Assets "InServiceYear" "In-Service Year"
F-Text   $L.Assets "YearSource" "Year Source"
F-Date   $L.Assets "PurchaseDate" "Purchase Date"
F-Date   $L.Assets "WarrantyEnd" "Warranty End"
F-Lookup $L.Assets "ReplacementClass" "Replacement Class" $L.Classes
F-Choice $L.Assets "BudgetPlan" "Budget Plan" @("Auto", "Replace", "Reserve", "Skip") "Auto"
F-Num    $L.Assets "PlanYear" "Plan Year"
F-Money  $L.Assets "QuotePrice" "Quoted Unit Price"
F-Text   $L.Assets "QuoteSource" "Quote Source"
F-Note   $L.Assets "AISuggestion" "AI Suggestion (JSON)"
F-Money  $L.Assets "MonthlyCost" "Monthly Cost"
F-Text   $L.Assets "ContractEnd" "Contract End"
F-Note   $L.Assets "Notes" "Notes"
F-Note   $L.Assets "ImportedAction" "Imported Action"
if ($new.Assets) { Set-View $L.Assets @("LinkTitle", "Category", "AssetType", "AssetLocation", "AssignedPerson", "Model", "SerialNumber", "AssetStatus", "InServiceYear", "ReplacementClass") }

$new.Actions = Ensure-List $L.Actions "Action"
F-Note   $L.Actions "Details" "Details"
F-Choice $L.Actions "Priority" "Priority" @("Critical", "High", "Medium", "Low") "Medium"
F-Lookup $L.Actions "ActionLocation" "Location" $L.Locations
F-Text   $L.Actions "Area" "Area"
F-Choice $L.Actions "ActionStatus" "Status" @("Open", "Done") "Open"
if ($new.Actions) { Set-View $L.Actions @("LinkTitle", "Priority", "ActionLocation", "Area", "ActionStatus") }

$new.Log = Ensure-List $L.Log "Asset Name"
if ($new.Log) { Set-PnPField -List $L.Log -Identity "Title" -Values @{ Required = $false } | Out-Null }
F-Lookup $L.Log "AssetRef" "Asset" $L.Assets -Index
F-Lookup $L.Log "PersonRef" "Person" $L.People -Index
F-Choice $L.Log "EntryType" "Entry Type" @("Note", "Repair", "Reimage", "Repurposed", "Assigned", "Unassigned", "Other") "Note"
F-Note   $L.Log "Details" "Details"
if ($new.Log) { Set-View $L.Log @("Created", "Author", "LinkTitle", "EntryType", "PersonRef", "Details") }

$new.Contracts = Ensure-List $L.Contracts "Contract"
F-Text   $L.Contracts "Vendor" "Vendor" -Index
F-Text   $L.Contracts "AccountNumber" "Account / Contract #"
F-Text   $L.Contracts "ServiceId" "Circuit / Service ID"
F-Choice $L.Contracts "ContractCategory" "Category" @("Connectivity", "Support Renewal", "Maintenance", "Software Subscription", "Licensing", "Service", "Lease", "Other") "Support Renewal"
F-Money  $L.Contracts "Cost" "Cost"
F-Choice $L.Contracts "Billing" "Billing" @("Monthly", "Annual", "Multi-year") "Annual"
F-Num    $L.Contracts "TermMonths" "Term (Months)"
F-Date   $L.Contracts "LastRenewal" "Last Renewal"
F-Date   $L.Contracts "NextRenewal" "Next Renewal" -Index
F-Bool   $L.Contracts "AutoRenew" "Auto-Renews"
F-Money  $L.Contracts "QuoteAmount" "Quote Amount"
F-Num    $L.Contracts "QuoteTermMonths" "Quote Term (Months)"
F-Text   $L.Contracts "QuoteNotes" "Quote Notes"
F-Lookup $L.Contracts "ContractLocation" "Location" $L.Locations
F-Text   $L.Contracts "Contact" "Contact"
F-Choice $L.Contracts "ContractStatus" "Status" @("Active", "Renewing", "Cancelled") "Active"
F-Note   $L.Contracts "ContractNotes" "Notes"
if ($new.Contracts) { Set-View $L.Contracts @("LinkTitle", "Vendor", "ContractCategory", "Cost", "Billing", "NextRenewal", "ContractLocation", "ContractStatus") }

# ---------------------------------------------------------------- starter data
function Count-Items($list) { @(Get-PnPListItem -List $list -PageSize 500).Count }

# replacement classes (only if the list is empty)
if ((Count-Items $L.Classes) -eq 0) {
  $classes = @(
    @("Business laptop", "laptop", "14-inch business laptop", 1400, 5, 2), @("Mobile workstation", "workstation", "15-inch mobile workstation", 2400, 5, 0),
    @("Desktop", "desktop", "Small-form-factor desktop", 950, 6, 1), @("Shop-floor tablet", "tablet", "Rugged tablet with mount", 700, 4, 0),
    @("Plant / OT computer", "ot", "Vendor-supplied machine PC", 2500, 10, 0), @("Access switch", "access_switch", "48-port PoE access switch", 4800, 10, 0),
    @("Core switch", "core_switch", "Stacked core switch pair", 18000, 12, 0), @("Wireless access point", "ap", "Wi-Fi 7 access point", 250, 7, 0),
    @("Firewall", "firewall", "Next-gen firewall", 2800, 6, 0), @("Virtualization host", "cluster_host", "2-socket rack server", 24000, 7, 0),
    @("Backup appliance", "backup", "Backup appliance", 7500, 5, 0), @("Industrial label printer", "label_printer", "4-inch industrial thermal printer", 2600, 8, 0),
    @("Office printer", "printer", "Business laser printer", 650, 6, 0), @("Multifunction copier", "mfp", "Department MFP", 5500, 6, 0),
    @("Handheld scanner", "scanner", "Rugged mobile computer", 1900, 5, 0), @("UPS", "ups", "Rack UPS 3kVA", 2100, 6, 0),
    @("IP phone", "phone", "Desk IP phone", 180, 8, 0), @("Not budgeted", "none", "", 0, 0, 0))
  foreach ($c in $classes) {
    Add-PnPListItem -List $L.Classes -Values @{ Title = $c[0]; ClassKey = $c[1]; StandardModel = $c[2]; UnitCost = [double]$c[3]; LifecycleYears = [int]$c[4]; SpareMin = [int]$c[5]; CostSource = "Example - set your own" } | Out-Null
  }
  Write-Host "Added $($classes.Count) replacement classes with example costs." -ForegroundColor Green
}

# locations from -Locations (skips names that already exist)
$locId = @{}; foreach ($i in Get-PnPListItem -List $L.Locations -PageSize 500) { $locId[("$($i['Title'])").ToLower()] = $i.Id }
foreach ($n in $Locations) {
  $k = $n.Trim().ToLower(); if (-not $k -or $locId.ContainsKey($k)) { continue }
  $locId[$k] = (Add-PnPListItem -List $L.Locations -Values @{ Title = $n.Trim() }).Id
  Write-Host "Added location: $n"
}

# fictional sample data
if ($SampleData) {
  if ((Count-Items $L.Assets) -gt 0 -or (Count-Items $L.People) -gt 0) { throw "-SampleData only loads into empty People and Assets lists. Use a new site or a different -Prefix." }
  if (-not (Test-Path $SampleFolder)) { throw "Sample data folder not found: $SampleFolder" }
  function Csv($f) { @(Import-Csv -Path (Join-Path $SampleFolder $f) -Encoding UTF8) }
  function V($x) { if ($null -eq $x -or "$x" -eq "") { $null } else { $x } }
  $classId = @{}; foreach ($i in Get-PnPListItem -List $L.Classes -PageSize 500) { $classId["$($i['ClassKey'])"] = $i.Id }
  foreach ($r in Csv "Locations.csv") { $k = $r.Title.ToLower(); if (-not $locId.ContainsKey($k)) { $locId[$k] = (Add-PnPListItem -List $L.Locations -Values @{ Title = $r.Title }).Id } }
  $personId = @{}
  foreach ($r in Csv "People.csv") {
    $v = @{ Title = $r.Name; PersonKey = $r.PersonKey; EmploymentStatus = $r.Status; O365 = ($r.O365 -eq "True"); MFA = ($r.MFA -eq "True"); VPN = ($r.VPN -eq "True"); AV = ($r.AV -eq "True") }
    if (V $r.JobTitle) { $v.JobTitle = $r.JobTitle }; if (V $r.Login) { $v.Login = $r.Login }
    if (V $r.Location) { $v.PersonLocation = $locId[$r.Location.ToLower()] }
    $personId[$r.PersonKey] = (Add-PnPListItem -List $L.People -Values $v).Id
  }
  Write-Host "Sample: $($personId.Count) people"
  $assetId = @{}; $n = 0
  foreach ($r in Csv "Assets.csv") {
    $v = @{ Title = $r.Name; AssetKey = ("a{0:d4}" -f ($n + 1)); Category = $r.Category; AssetStatus = $r.Status; BudgetPlan = $(if (V $r.BudgetPlan) { $r.BudgetPlan } else { "Auto" }); Quantity = [int]$(if (V $r.Quantity) { $r.Quantity } else { 1 }) }
    foreach ($k in "AssetType", "SharedUse", "Make", "Model", "SerialNumber", "IPAddress", "Setup", "YearSource", "QuoteSource", "Notes") { if (V $r.$k) { $v[$k] = $r.$k } }
    if (V $r.InServiceYear) { $v.InServiceYear = [int]$r.InServiceYear }; if (V $r.PlanYear) { $v.PlanYear = [int]$r.PlanYear }
    if (V $r.QuotePrice) { $v.QuotePrice = [double]$r.QuotePrice }
    if (V $r.PurchaseDate) { $v.PurchaseDate = $r.PurchaseDate }; if (V $r.WarrantyEnd) { $v.WarrantyEnd = $r.WarrantyEnd }
    if (V $r.Location) { $v.AssetLocation = $locId[$r.Location.ToLower()] }
    if ((V $r.PersonKey) -and $personId.ContainsKey($r.PersonKey)) { $v.AssignedPerson = $personId[$r.PersonKey] }
    if ((V $r.ClassKey) -and $classId.ContainsKey($r.ClassKey)) { $v.ReplacementClass = $classId[$r.ClassKey] }
    $assetId[$r.Name] = (Add-PnPListItem -List $L.Assets -Values $v).Id
    $n++; if ($n % 50 -eq 0) { Write-Host "  $n assets..." }
  }
  Write-Host "Sample: $n assets"
  foreach ($r in Csv "Actions.csv") {
    $v = @{ Title = $r.Title; Details = $r.Details; Priority = $r.Priority; Area = $r.Area; ActionStatus = $r.Status }
    if (V $r.Location) { $v.ActionLocation = $locId[$r.Location.ToLower()] }
    Add-PnPListItem -List $L.Actions -Values $v | Out-Null
  }
  foreach ($r in Csv "Contracts.csv") {
    $v = @{ Title = $r.Title; Vendor = $r.Vendor; ContractCategory = $r.Category; Cost = [double]$r.Cost; Billing = $r.Billing; TermMonths = [int]$r.TermMonths
            NextRenewal = $r.NextRenewal; AutoRenew = ($r.AutoRenew -eq "True"); ContractStatus = "Active" }
    foreach ($k in "AccountNumber", "ServiceId", "LastRenewal", "QuoteNotes", "Contact") { if (V $r.$k) { $v[$k] = $r.$k } }
    if (V $r.QuoteAmount) { $v.QuoteAmount = [double]$r.QuoteAmount; $v.QuoteTermMonths = [int]$r.QuoteTermMonths }
    if (V $r.Location) { $v.ContractLocation = $locId[$r.Location.ToLower()] }
    Add-PnPListItem -List $L.Contracts -Values $v | Out-Null
  }
  foreach ($r in Csv "AssetLog.csv") {
    if (-not $assetId.ContainsKey($r.Asset)) { continue }
    Add-PnPListItem -List $L.Log -Values @{ Title = $r.Asset; AssetRef = $assetId[$r.Asset]; EntryType = $r.EntryType; Details = $r.Details } | Out-Null
  }
  Write-Host "Sample data loaded (fictional company: Contoso Manufacturing)." -ForegroundColor Green
}

Write-Host ""
Write-Host "Done. Lists are ready on $SiteUrl" -ForegroundColor Green
Write-Host "Next: upload it-asset-tracker.sppkg to your App Catalog, add the app to this site, and add the 'IT Asset Tracker' web part to a page."
