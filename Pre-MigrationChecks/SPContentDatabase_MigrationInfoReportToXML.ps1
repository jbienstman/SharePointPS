<#
.NOTES
    ########################################################################################################################################
    # Author: Jim B.
    ########################################################################################################################################
    # Revision(s)
    # 1.2.0  Added region - "User Policy Settings"
    # 1.3.0  Changed $allDbs = Get-SPContentDatabase --> $allDbs = Get-SPDatabase | Where-Object {($_.Type -eq "Content Database") -and ($_.WebApplication -notlike "*Central*")}
    #        Added "IF" clause to region - "User Policy Settings", default = NOT EXECUTED
    # 1.4.0  Compensated for possible Site Collections without a Template and No RootWeb/AllWebs output
    #        Changed HEADER to include all checks - exclude "General Header Functions"
    #        Write XML in UTF-8
    # 1.4.1  2018-02-05 - Updated Header Section, cleanup REGIONS, cleanup Web App Policy Section
    # 1.4.2  2018-02-06 - Changed Test-SPContent Database to use more distinguishing properties: -Name $db.DisplayName -WebApplication $db.WebApplication.Url -ServerInstance $db.Server
    #                     Noticed that not providing the "ServerInstance" can block the execution if there are multiple aliases defined
    # 1.4.5  2018-02-07 - Updated Header & Try/Catch
    # 1.5.0  2018-10-09 - NOTE: Need to include large list check
    # 1.5.2  2019-07-19 - Updated User Policy to Function, using proper XML constructor now...
    # 1.6.0  2019-07-29 - Added Large List Check, changed XML to include Web Application & Updated Header and cleanup of various typos and stuff
    # 1.6.1  2021-04-06 - Added condition to detect whether a database actually contains sites before doing site related tests
    #
    ########################################################################################################################################
.SYNOPSIS
    Scans SharePoint 2013 & 2016 Databases for potential "Database Attach" Migration Issues and reports these to XML.
    NOTE: Although the script has been tested successfully at multiple clients, there are NO GUARANTEES THAT THIS WILL WORK ON ANY FARM!
.DESCRIPTION
    This script will iterate through all webapplications and all content databases (*wildcard exclusion possible for both*) and Site 
    Collections and Webs to check for the most common Migration Issues such as: Backwards Compatibility of Databases and/or Site collections,
    Errors (Missing Dependenies), Large Lists and outputs a findings to XML in the PSSCriptRoot: databaseMigrationReport_yyyyMMddHHmmss.xml.
    NOTE: The script will put some load on your Farm, so it is recommended to run this outside of business hours if this would pose a problem.
    NOTE: When used on SharePoint 2016, site collections are all still in 15 mode by default --> will trigger the "HasCompatabilityModeSPSite"
    
.LINK
    ...
.EXAMPLE
    SPContentDatabase_MigrationInfoReportToXML.ps1

    For Example - on a SharePoint 2016 Single Server with 8 Web Applications, 33 Databases, 36 Sites and 92 Webs 
    and 11 Large Lists, the script took under 107 seconds to complete with 38K XML file, additional data collections took about 1/2 that time.
    There are other related scripts that will help with fixing the mentioned Issues.
    No Parameters needed!
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
#region - SharePoint Functions
Function Set-WebAppsUserPolicy {
    <#
    .SYNOPSIS
        NOTE: Requires Farm Administrator Permissions and Web Application Policy FullRead for iteration through the Site Collections
        Iterates through all Web Applications and sets user policy to "FullControl" or "FullRead" permissions for a single user (Format: DOMAIN\LOGIN) on all zones. 
        There is an optional ObjectCacheTag Parameter that allows you to also add the required property to the web application.
    .EXAMPLE
        Set-WebAppsUserPolicy -account DOMAIN\LOGIN -SPPolicyRoleType FullControl -ObjectCacheTag portalsuperuseraccount
    #>
    Param (
        [Parameter(Mandatory=$true)][string]$account ,
        [Parameter(Mandatory=$true)][ValidateSet("FullControl","FullRead")][string]$SPPolicyRoleType ,
        [Parameter(Mandatory=$false)][ValidateSet("portalsuperuseraccount","portalsuperreaderaccount")][string]$ObjectCacheTag
    )     
    write-host ("Setting `"$SPPolicyRoleType`" for `"$account`" on User Policy of all Web Applications:") -ForegroundColor Cyan
    $webApps = Get-SPWebApplication
    foreach($webApp in $webApps)
    {
    #region - classic or claims check
    if ($webApp.UseClaimsAuthentication -eq $True)
        {
        $setAccount = (New-SPClaimsPrincipal -identity $account -identitytype 1).ToEncodedString()
        }
    else
        {
        $setAccount = $account
        }
    #endregion
    if ($ObjectCacheTag -eq "portalsuperuseraccount")
        {
        write-host (" Setting portalsuperUSERaccount on: " + $webApp.DisplayName) -NoNewline
        $webApp.Properties["portalsuperuseraccount"] = $setAccount
        write-host ("...Done") -ForegroundColor Green
        }
    elseif($ObjectCacheTag -eq "portalsuperreaderaccount")
        {
        write-host (" Setting portalsuperREADERaccount on: " + $webApp.DisplayName) -NoNewline                
        $webApp.Properties["portalsuperreaderaccount"] = $setAccount
        write-host ("...Done") -ForegroundColor Green
        }
    else
        {
        }
    write-host (" Setting `"" + $SPPolicyRoleType + "`" on: `"" + $webApp.DisplayName + "`" for: `"" + $setAccount + "`"") -NoNewline            
    $policy = $webApp.Policies.Add($setAccount,$setAccount)
    $policyRole = $webApp.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::($SPPolicyRoleType))
    $policy.PolicyRoleBindings.Add($policyRole)
    $webApp.Update()
    write-host "...Done" -ForegroundColor Green
    }
    write-host ("Finished") -ForegroundColor Cyan
}
#endregion
########################################################################################################################################
#region - MAIN SCRIPT VARIABLES
#
$scriptAction = "Content Database Migration Info Report"
$xmlFilePathNamePrefix = "databaseMigrationReport"
$startDateTime = Get-date -Format yyyyMMddHHmmss
$xmlFilePath = ($ScriptPath + "\" + $xmlFilePathNamePrefix + "_" + $startDateTime + ".xml")
$webApplicationExclusionString = "*WILDCARD_NAME_OF_WEBAPP_YOU_WANT_TO_EXCLUDE_HERE*"
$databaseNameExclusionString = "*WILDCARD_NAME_OF_DATABASE_YOU_WANT_TO_EXCLUDE_HERE*"
#
$setUserPolicySettings = $false
$UserPolicyPermissionLevel = "FullControl" #FullControl FullRead
#
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
#region - User Policy Settings
#
if ($setUserPolicySettings -eq $true)
{
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name  
    Set-WebAppsUserPolicy -account $currentUser -SPPolicyRoleType $UserPolicyPermissionLevel
}
#
#endregion - User Policy Settings
#region - Farm Iteration
# NOTE: The Get-SPContentDatabase command doesn't return "offline" databases, you can use Get-SPDatabase instead (includes the NeedsUpgrade tags)
#region - Getting Top Level Stuff
$farm = Get-SPFarm
$webApps = Get-SPWebApplication
$allDbs = Get-SPDatabase | Where-Object {($_.Type -eq "Content Database") -and ($_.WebApplication -notlike "*SPAdministrationWebApplication*") -and ($_.WebApplication -notlike $webApplicationExclusionString)}
$allFilteredDatabases = $allDbs | Where-Object {$_.Name -notlike $databaseNameExclusionString}
Write-Host ("Content Databases to Iterate: " ) -NoNewline
Write-Host $allFilteredDatabases.Count -ForegroundColor Cyan -NoNewline
Write-Host (" - Excluding: ") -NoNewline
Write-Host ($allDbs.count - $allFilteredDatabases.Count) -NoNewline -ForegroundColor Gray
Write-Host (" (Filter: ") -NoNewline
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
        $null = $xmlDatabaseElement
        $xmlDatabaseElement.SetAttribute("DisplayName",$db.DisplayName)
        $xmlDatabaseElement.SetAttribute("Sites",$db.CurrentSiteCount)
        if ($db.CurrentSiteCount -gt 0)
            {
            #region - Get Orphans
            [xml]$orphanedObjects = $db.Repair($false)
            $NumberOfOrphans = $orphanedObjects.OrphanedObjects.Count    
            #endregion - Get Orphans
            #region - Get dbErrors
            $dbErrors = Test-SPContentDatabase -Name $db.DisplayName -WebApplication $db.WebApplication.Url -ServerInstance $db.Server
            #endregion - Get dbErrors
            Write-Host (" Checking ") -NoNewline
            If ($farm.BuildVersion.Major -lt 16)
                {
                if (Get-SPSite -ContentDatabase $db.DisplayName -Limit All -ErrorAction SilentlyContinue -WarningAction SilentlyContinue| Where-Object {$_.CompatibilityLevel -ne $farm.BuildVersion.Major.ToString()})    
                    {                        
                    Write-Host " HasCompatabilityModeSPSite: " -NoNewline
                    Write-Host "True" -NoNewline -ForegroundColor Yellow
                    $xmlDatabaseElement.SetAttribute("HasCompatabilityModeSPSite","True")
                    }
                else
                    {        
                    Write-Host " HasCompatabilityModeSPSite: " -NoNewline
                    Write-Host "False" -NoNewline -ForegroundColor Cyan
                    $xmlDatabaseElement.SetAttribute("HasCompatabilityModeSPSite","False")
                    }
                }
            else
                {
                #No Compatibility Mode exists after SharePoint 2013
                }
            $xmlDatabaseElement.SetAttribute("NeedsUpgrade",$db.NeedsUpgrade)
            $xmlDatabaseElement.SetAttribute("NeedsUpgradeIncludeChildren",$db.NeedsUpgradeIncludeChildren)
            if (Get-SPSite -ContentDatabase $db.DisplayName -Limit all -ErrorAction SilentlyContinue  -WarningAction SilentlyContinue| Where-Object {$_.NeedsUpgrade -eq $true})
                {
                $xmlDatabaseElement.SetAttribute("SPSiteWithNeedsUpgrade","True")
                }
            else
                {
                $xmlDatabaseElement.SetAttribute("SPSiteWithNeedsUpgrade","False")
                }
            $xmlDatabaseElement.SetAttribute("Server",$db.Server)
            #region - Orphan Output
            if ($NumberOfOrphans -gt "0")
                {   
                $OrphanTypes = ""
                Write-Host " - Orphans: " -NoNewline
                Write-Host $NumberOfOrphans -NoNewline -ForegroundColor Cyan        
                Write-Host " (" -NoNewline        
                foreach ($orphanType in $orphanedObjects.OrphanedObjects.ChildNodes)
                    {
                    $oType = $orphanType.Type.ToString()
                    $OrphanTypes += ($oType + ",")
                    Write-Host ":$oType " -NoNewline -ForegroundColor Yellow            
                    }
                #region - xmlStringAddAttribute
                $xmlDatabaseElement.SetAttribute("Orphans",$NumberOfOrphans)
                $xmlDatabaseElement.SetAttribute("OrphanTypes",$OrphanTypes.Trim(","))
                #endregion - xmlStringAddAttribute
                Write-Host ")" -NoNewline  
                }
            else
                {        
                Write-Host " - Orphans: " -NoNewline
                Write-Host "None ()" -NoNewline -ForegroundColor Green
                $xmlDatabaseElement.SetAttribute("Orphans",0)
                }
            #endregion - Orphan Output
            #region - dbErrors Output
            if ($null -eq $dbErrors)
                {        
                Write-Host " - Errors: " -NoNewline
                write-Host "No " -ForegroundColor Green -NoNewline
                #region - xmlStringAddAttribute
                $xmlDatabaseElement.SetAttribute("Errors","False")
                #endregion - xmlStringAddAttribute
                }
            else
                {        
                Write-Host " - Errors: " -NoNewline
                write-Host "Yes " -ForegroundColor Cyan -NoNewline
                #region - xmlStringAddAttribute
                $xmlDatabaseElement.SetAttribute("Errors","True")
                #endregion - xmlStringAddAttribute
                }    
            #endregion - dbErrors Output
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
                $xmlSiteCollectionElement = $xmlFile.CreateElement("SiteCollection")
                $xmlSiteCollectionElement.SetAttribute("Url",$spSite.Url)
                $xmlSiteCollectionElement.SetAttribute("CompatibilityLevel",$spSite.CompatibilityLevel)
                $xmlSiteCollectionElement.SetAttribute("NeedsUpgrade",$spSite.NeedsUpgrade)
                $rootWeb = Get-SPWeb -Site $spSite.Url -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                        if ($rootWeb.WebTemplate)
                            {
                            $webCount += $spSite.AllWebs.Count
                            $xmlSiteCollectionElement.SetAttribute("AllWebs",$spSite.AllWebs.Count.ToString())
                            $spSiteHasRootWeb = $true
                            }
                        else
                            {
                            $xmlSiteCollectionElement.SetAttribute("AllWebs",0)
                            }
                $xmlSiteCollectionElement.SetAttribute("RecycleBinWebs",($spSite.RecycleBin | Where-Object {$_.ItemType -like "web"}).count)
                $xmlSiteCollectionElement.SetAttribute("RecycleBinAllItems",$spSite.RecycleBin.Count)
                $xmlSiteCollectionElement.SetAttribute("HostHeaderIsSiteName",$spSite.HostHeaderIsSiteName)
                $xmlSiteCollectionElement.SetAttribute("SchemaVersion",$spSite.SchemaVersion.ToString())
                $xmlDatabaseElement.AppendChild($xmlSiteCollectionElement) | Out-Null
                $xmlFile.Save($xmlFilePath)
                if ($spSiteHasRootWeb)
                    {
                    #region - Xml SPWeb                
                    foreach($SPweb in $SPSite.AllWebs) #SPSite.AllWebs lists ALL Subwebs, no matter how deep the path goes
                                                                                                                                                                                    {
                    $webHasLargeList = $false
                    $null = $xmlSPWebElement
                    $xmlSPWebElement = $xmlFile.CreateElement("spweb")
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
                            $xmlSPListElement = $xmlFile.CreateElement("splist")
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
        else
            {
            Write-Host (" No sites in this database") -ForegroundColor Yellow
            $xmlSPWebApplicationElement.AppendChild($xmlDatabaseElement) | Out-Null
            $xmlFile.Save($xmlFilePath)
            }
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
Write-Host ("You can find the XML report here: `"" + $xmlFilePath + "`"") -ForegroundColor Yellow
#endregion - Farm Iteration
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
