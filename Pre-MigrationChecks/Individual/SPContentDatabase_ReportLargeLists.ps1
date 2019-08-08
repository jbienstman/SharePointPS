<#
.NOTES
    ########################################################################################################################################
    # Author: Jim B.
    ########################################################################################################################################
    # Revision(s)
    # 1.2.0 - 2019-08-08 - Initial Commit to GitHub
    # 1.3.0 - 2019-08-08 - Updated to also report the Large Lists and create a separate CSV output, fixed exlusion via Web Application
    ########################################################################################################################################
.SYNOPSIS
    Scans SharePoint 2013 & 2016 Databases for Large Lists and reports then in an xml 
.DESCRIPTION
    Scans SharePoint 2013 & 2016 Databases for Large Lists and reports then in an xml    
.LINK
    ...
.EXAMPLE
    ...
#>
###########################################################################################################################################
#region - TRY
try 
{
########################################################################################################################################
#region - Top Level Script Settings
$startTime = (Get-Date)
Clear-Host
Write-Host ""
Write-Host ("Script Started: ") -NoNewline -ForegroundColor DarkCyan
Write-Host ((Get-Date -Format F)) -ForegroundColor DarkYellow
Write-Host ""
$ErrorActionPreference = "Stop";
#endregion - Top Level Script Settings
########################################################################################################################################
#region - HEADER v1.2.7
#region - ChangeLog
# v1.0.0 - 2018-01-01 - Various Iterations
# v1.1.0 - 2018-02-05 - Started Change Log - Removed Date & Time Section
# v1.2.0 - 2018-02-07 - Updated the Inclusions Section, moved variable up, display on screen in better colors
# v1.2.1 - 2018-06-26 - Removed Start-SPAssignment section fom Header - Moved to MAIN
# v1.2.2 - 2018-07-02 - Added Transcript Option, added Start-SPAssignment again, but with fault handling
# v1.2.3 - 2018-07-10 - Added Write-Progress Example, requires an acion of at least 1 second before it gets displayed & $PSScriptRoot can replace the majority of the "getting script path region"
# v1.2.4 - 2018-08-20 - Cleanup & Formatting of the Header Script
# v1.2.5 - 2018-11-05 - Make Run as administrator & Loading SharePoint Snap-In Optional (removed Write-Progress part also)
# v1.2.6 - 2018-12-13 - Changed "ScriptPath" section, now this part always executes, as it is a built-in parameter anyway...changed some comments and OK! to OK
# v1.2.7 - 2019-07-15 - Updated Output colors and swithed position of a $null compare
#endregion - ChangeLog
#region - HEADER VARIABLES
$InclusionScriptNames = @("")
$ScriptRequiresAdminPrivileges = $False
$ScriptRequiresSharePointAddin = $True
#OVERRIDE: SHOULD BE REPLACED BY PARAMETER IN SCRIPT TEMPLATE
$EnableLogging = $False
#
#endregion - HEADER VARIABLES
#region - Get Script Path
Write-Host "Getting ScriptPath: " -NoNewline -ForegroundColor Gray
$ScriptPath = $PSScriptRoot #This Should ALWAYS EXECUTE
if ($MyInvocation.InvocationName.Length -ne "0")
    {    
    #$ScriptPath = Split-Path $MyInvocation.InvocationName
    $ScriptFileName = $MyInvocation.MyCommand.Name           
    $ScriptFileNameNoExtension = ($ScriptFileName.Split("."))[0]
    Write-Host "OK" -ForegroundColor Green        
    }
else 
    {
    Write-Host "Cannot get Script Path, stopping script..." -ForegroundColor Yellow
    exit
    }
#endregion - Get Script Path
#region - Inclusion(s)
Write-Host ("Including Script(s): ") -NoNewline -ForegroundColor Gray
Write-Host ($InclusionScriptNames) -NoNewline -ForegroundColor DarkGray
if ($InclusionScriptNames)
    {
    foreach ($InclusionScriptName in $InclusionScriptNames)
        {        
        $InclusionScriptFullPath = ($ScriptPath + "\"  + $InclusionScriptName)
        if (Test-Path $InclusionScriptFullPath)
            {
            . $InclusionScriptFullPath
            Write-Host "...Done" -ForegroundColor Green
            }
        else
            {
            Write-Host "...Error" -ForegroundColor Red
            Write-Host ($InclusionScriptFullPath + " - NOT FOUND") -ForegroundColor Yellow
            exit
            }        
        }
    }
else
    {
    Write-Host "0" -ForegroundColor Green
    }
#endregion - Inclusion(s)
#region - Run As Admin
if ($ScriptRequiresAdminPrivileges)
    {
    Function IsAdmin 
        {
        $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")    
        return $IsAdmin
        }
    Write-Host "Run As Admin: " -NoNewline -ForegroundColor Gray
    if((IsAdmin) -eq $false)
        {
        Write-Host "NO" -ForegroundColor Red
        Write-Warning "This Script requires `"Administrator`" privileges, stopping..."
        return
        }
    else
        {
        Write-Host "OK" -ForegroundColor Green
        }    
    }
#endregion - Run As Admin
#region - SharePoint Snapin
if ($ScriptRequiresSharePointAddin)
    {    
    Write-Host "Loading SharePoint PowerShell Snapin: " -NoNewline -ForegroundColor Gray
    if ($null -eq (Get-PSSnapin "Microsoft.SharePoint.PowerShell" -WarningAction SilentlyContinue -ErrorAction SilentlyContinue))
        {
        Add-PSSnapin "Microsoft.SharePoint.PowerShell" -WarningAction SilentlyContinue -ErrorAction Stop
        $SharePointPowershellSnapinLoaded = $true
        Write-Host "OK" -ForegroundColor Green        
        }
    else
        {
        Write-Host "OK (Already Loaded)" -ForegroundColor Green
        $SharePointPowershellSnapinLoaded = $true
        }    
    #region - Start-SPAssignment
    if ($SharePointPowershellSnapinLoaded) {$StartSPAssignMentSwitchedOn = Start-SPAssignment -Global -WarningAction SilentlyContinue -ErrorAction SilentlyContinue}
    #endregion - Start-SPAssignment
    }
#endregion - SharePoint Snapin
#region - Transcript
if ($EnableLogging) {Start-Transcript -Path ($ScriptPath + "\" + $ScriptFileNameNoExtension + ".log")}
#endregion - Transcript
#endregion - HEADER
###########################################################################################################################################
#region - MAIN SCRIPT VARIABLES
#
$scriptAction = "Content Database Large List Report"
$xmlFilePathNamePrefix = "databaseLargeListsReport"
$startDateTime = Get-date -Format yyyyMMddHHmmss
$xmlFilePath = ($ScriptPath + "\" + $xmlFilePathNamePrefix + "_" + $startDateTime + ".xml")
$webApplicationExclusionString = "*WILDCARD_NAME_OF_WEBAPP_YOU_WANT_TO_EXCLUDE_HERE*"
$databaseNameExclusionString = "*WILDCARD_NAME_OF_DATABASE_YOU_WANT_TO_EXCLUDE_HERE*"
[int]$itemCountThreshold = 5000 #default = 5000
#
#endregion - MAIN SCRIPT VARIABLES
########################################################################################################################################
#region - Output Script Title
Write-Host
Write-Host $scriptAction -ForegroundColor Cyan
for ($i = 0; $i -lt $scriptAction.Length; $i++)
    {
    Write-Host "=" -NoNewline -ForegroundColor Cyan
    }
Write-Host
#endregion - Output Script Title
#region - Farm Iteration
# NOTE: The Get-SPContentDatabase command doesn't return "offline" databases, you can use Get-SPDatabase instead (includes the NeedsUpgrade tags)
#region - Getting Top Level Stuff
$farm = Get-SPFarm
$webApps = Get-SPWebApplication
$allDbs = Get-SPDatabase | Where-Object {($_.Type -eq "Content Database") -and ($_.WebApplication -notlike "*SPAdministrationWebApplication*")}
$allDbsFilteredWebApp = Get-SPDatabase | Where-Object {($_.Type -eq "Content Database") -and ($_.WebApplication -notlike "*SPAdministrationWebApplication*") -and ($_.WebApplication -notlike $webApplicationExclusionString)}
$allFilteredDatabases = $allDbsFilteredWebApp | Where-Object {$_.Name -notlike $databaseNameExclusionString}
Write-Host ("Content Databases to Iterate: " ) -NoNewline
Write-Host $allFilteredDatabases.Count -ForegroundColor Cyan -NoNewline
Write-Host (" - Excluding: ") -NoNewline
Write-Host ($allDbs.count - $allFilteredDatabases.Count) -NoNewline -ForegroundColor Yellow
Write-Host (" (FilterWebApp: ") -NoNewline
Write-Host $webApplicationExclusionString -ForegroundColor Gray -NoNewline
Write-Host (" | FilterDB: ") -NoNewline
Write-Host $databaseNameExclusionString -ForegroundColor Gray -NoNewline
Write-Host (") databases from iteration...")
Write-Host
[int]$longestDbName = 0
foreach ($db in $allFilteredDatabases)
    {
    if ($db.DisplayName.Length -gt $longestDbName)
        {
        $longestDbName = $db.DisplayName.Length
        }
    }
$databaseCounterDigits = $allFilteredDatabases.Count.ToString().Length
#endregion - Getting Top Level Stuff
#region - Xml Root
$xmlFile = New-Object Xml.XmlDocument
$xmlDeclaration = $xmlFile.CreateXmlDeclaration("1.0","utf-8",$null)
$xmlFile.InsertBefore($xmlDeclaration, $xmlFile.DocumentElement) | Out-Null
$xmlRootElement = $xmlFile.CreateElement("farm")
$xmlRootElement.SetAttribute("Name",$farm.Name)
$xmlRootElement.SetAttribute("BuildVersion",$farm.BuildVersion.ToString())
$xmlRootElement.SetAttribute("MajorVersion",$farm.BuildVersion.Major.ToString())
$xmlRootElement.SetAttribute("WebApplications",$webApps.Count)
$xmlRootElement.SetAttribute("AllSPContentDatabases",$allDbs.Count)
$xmlRootElement.SetAttribute("ExclusionFilter",$databaseNameExclusionString)
$xmlRootElement.SetAttribute("FilteredSPContentDatabases",$allFilteredDatabases.Count)
$xmlFile.AppendChild($xmlRootElement) | Out-Null
$xmlFile.Save($xmlFilePath)
#endregion - Xml Root
#region - Xml WebApp
$databaseCounterIndex = 1
$farmSiteCount = 0
$farmWebCount = 0
$farmLargeListCount = 0
foreach ($SPwebApp in $webApps)
    {
    $webCount = 0
    Write-Host ("WebApplication: `"" + $SPwebApp.DisplayName + "`"") -ForegroundColor DarkGray    
    $null = $xmlSPWebApplicationElement
    $xmlSPWebApplicationElement = $xmlFile.CreateElement("webapplication")
    $xmlSPWebApplicationElement.SetAttribute("DisplayName",$SPwebApp.DisplayName)
    $xmlSPWebApplicationElement.SetAttribute("Url",$SPwebApp.Url)
    $xmlSPWebApplicationElement.SetAttribute("ContentDatabases",$SPwebApp.ContentDatabases.Count)
    $xmlSPWebApplicationElement.SetAttribute("Sites",$SPwebApp.Sites.Count)
    $xmlRootElement.AppendChild($xmlSPWebApplicationElement) | Out-Null
    $xmlFile.Save($xmlFilePath)
#endregion - Xml WebApp
    #region - DATABASE ITERATION
    foreach ($db in $allFilteredDatabases | Where-Object {$_.WebApplication.Name -eq $SPwebApp.Name})
        {
        #region - Formatted String Output
        [int]$spacesDbName = (($longestDbName + 2) - $db.DisplayName.Length)
        Write-Host (" - [{0:d$databaseCounterDigits}/" -f $databaseCounterIndex) -NoNewline -ForegroundColor Gray ;
        Write-Host ($allFilteredDatabases.Count.ToString() + "] ") -NoNewline -ForegroundColor Gray
        Write-Host ("{0}{1,$spacesDbName}" -f $db.DisplayName, ":") -NoNewline ;
        $databaseCounterIndex++
        #endregion - Formatted String Output
        #region - Xml Database 
        $xmlDatabaseElement = $xmlFile.CreateElement("Database")
        $xmlDatabaseElement.SetAttribute("Server",$db.Server)
        $xmlSPWebApplicationElement.AppendChild($xmlDatabaseElement) | Out-Null
        $xmlFile.Save($xmlFilePath)
        #endregion - Xml Database 
        #region - Xml Site Collection
        $spSites = Get-SPSite -ContentDatabase $db.DisplayName -Limit all -ErrorAction SilentlyContinue  -WarningAction SilentlyContinue
        write-Host "(Iterating through " -NoNewline
        Write-Host $spSites.Count -ForegroundColor Cyan -NoNewline
        Write-Host " Site Collections)" -NoNewline
        foreach ($spSite in $spSites)
            {
            $spSiteHasRootWeb = $false
            $null = $xmlSiteCollectionElement
            $xmlSiteCollectionElement = $xmlFile.CreateElement("SPSite")
            $xmlSiteCollectionElement.SetAttribute("Url",$spSite.Url)
            $rootWeb = Get-SPWeb -Site $spSite.Url -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                    if ($rootWeb.WebTemplate)
                        {
                        $webCount += $spSite.AllWebs.Count
                        $spSiteHasRootWeb = $true
                        }
                    else
                        {
                        }
            $xmlDatabaseElement.AppendChild($xmlSiteCollectionElement) | Out-Null
            $xmlFile.Save($xmlFilePath)
            if ($spSiteHasRootWeb)
                {
                #region - Xml SPWeb                
                foreach($SPweb in $SPSite.AllWebs) #SPSite.AllWebs lists ALL Subwebs, no matter how deep the path goes
                    {
                    $webHasLargeList = $false
                    $null = $xmlSPWebElement
                    $xmlSPWebElement = $xmlFile.CreateElement("SPWeb")
                    $xmlSPWebElement.SetAttribute("Url",$SPweb.Url)
                    $xmlSPWebElement.SetAttribute("IsRootWeb",$SPweb.IsRootWeb)
                    $xmlSPWebElement.SetAttribute("HasUniquePerm",$SPweb.HasUniquePerm)
                    $xmlSPWebElement.SetAttribute("HasUniqueRoleAssignments",$SPweb.HasUniqueRoleAssignments)
                    $xmlSPWebElement.SetAttribute("HasUniqueRoleDefinitions",$SPweb.HasUniqueRoleDefinitions)                    
                    $xmlSiteCollectionElement.AppendChild($xmlSPWebElement) | Out-Null
                    $xmlFile.Save($xmlFilePath)
                    #region - Xml SPList
                    foreach($SPlist in $SPweb.Lists) 
                        {
                        if($splist.ItemCount -gt $itemCountThreshold) 
                            {
                            $farmLargeListCount++
                            $webHasLargeList = $true
                            $null = $xmlSPListElement
                            $xmlSPListElement = $xmlFile.CreateElement("SPList")
                            $xmlSPListElement.SetAttribute("Title",$SPlist.Title)
                            $xmlSPListElement.SetAttribute("Url",($SPwebApp.Url.TrimEnd("/") + $SPlist.DefaultViewUrl))
                            $xmlSPListElement.SetAttribute("Items",$SPlist.ItemCount)
                            $xmlSPWebElement.AppendChild($xmlSPListElement) | Out-Null
                            $xmlFile.Save($xmlFilePath) 
                            }
                        }
                    #endregion - Xml SPList
                    if ($webHasLargeList)
                        {
                        $xmlSPWebElement.SetAttribute("HasLargeList","True")
                        $xmlFile.Save($xmlFilePath)
                        }
                    else
                        {
                        $xmlSPWebElement.SetAttribute("HasLargeList","False")
                        $xmlFile.Save($xmlFilePath)
                        }
                    $SPweb.Dispose()
                    }
                #endregion - Xml SPWeb
                }
            }
            $SPsite.Dispose()
        write-Host "...Done" -ForegroundColor Green
        #endregion - Xml Site Collection
        }
    #endregion - DATABASE ITERATION
    $xmlSPWebApplicationElement.SetAttribute("Webs",$webCount)
    $xmlFile.Save($xmlFilePath)
    $farmSiteCount += $SPwebApp.Sites.Count
    $farmWebCount += $webCount    
    }
    $xmlRootElement.SetAttribute("Sites",$farmSiteCount)
    $xmlRootElement.SetAttribute("Webs",$farmWebCount)
    $xmlRootElement.SetAttribute("LargeLists",$farmLargeListCount)
    $xmlFile.Save($xmlFilePath)
Write-Host
Write-Host ("You can find the XML report here: `"" + $xmlFilePath + "`"") -ForegroundColor Gray
#endregion - Farm Iteration
#region - Retrieve Input XML
[xml]$XMLFile = Get-Content $xmlFilePath
#endregion - Retrieve Input XML
#region - parsing
$largeLists = $XMLFile.farm.webapplication.database.spsite.spweb.splist
$largeListsCount = ($largeLists | Measure-Object).Count
if ($largeListsCount -gt 0)
    {
    [array]$ArrayOfListObjects = @()
    Write-Host ("Found " + $largeListsCount + " Large Lists - Parsing") -ForegroundColor Yellow -NoNewline
    foreach ($splist in $largeLists)
        {    
        $ArrayOfListObjects += $splist    
        }
    Write-Host " ...Done" -ForegroundColor Green
    #region - output    
    $csvFilePath = ($xmlFilePath.TrimEnd(".xml") + ".csv")
    $ArrayOfListObjects | Select-Object Title, Url, Items | Export-Csv -Path $csvFilePath -Encoding UTF8
    Write-Host ("You can find the CSV report with the Large Lists here: `"" + $csvFilePath + "`"") -ForegroundColor Gray
    #endregion - output
    }
else
    {
    Write-Host ("NO Lists with more than `"" + $itemCountThreshold + "`" items was found.") -ForegroundColor Green
    }
#endregion - parsing


########################################################################################################################################
}
#endregion - TRY
########################################################################################################################################
#region - Catch & Finally
catch
{
    #region - Catch
    Write-Host "|------------------> <Catch - Error Description>`n" -ForegroundColor Yellow
    Write-Host $Error[0].Exception -ForegroundColor Red
    Write-Host "`n<------------------| </Catch - Error Description>" -ForegroundColor Yellow;
    #endregion - Catch
}
finally
{
    #region - Finally
    if ($null -ne $StartSPAssignMentSwitchedOn)
        {
        Stop-SPAssignment -Global | Out-Null   
        }
    $ErrorActionPreference = "Continue";
    Write-Host ""
    $endTime = (Get-Date)
    Write-Host ("Script Ended: ") -NoNewline -ForegroundColor Cyan
    Write-Host ((Get-Date -Format F)) -ForegroundColor DarkYellow -NoNewline      
    Write-Host (" (and took: ") -ForegroundColor Gray -NoNewline
    $timeSpan = ($($endTime - $startTime).TotalSeconds)
    Write-Host ([math]::Round($timeSpan,3)) -NoNewline
    Write-Host (" Seconds)") -ForegroundColor Gray
    #endregion - Finally
}
#endregion - Catch & Finally
########################################################################################################################################