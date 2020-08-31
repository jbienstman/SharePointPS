<#
.NOTES
    ########################################################################################################################################
    # Author: Jim B.
    ########################################################################################################################################
    # Revision(s)
    # 1.0.0 - ????-??-?? - Don't remember...
    # 1.3.1 - 2017-11-03 - Compensated for possible Site Collections without a Template and No RootWeb/AllWebs output
    # 1.3.2 - 2017-11-07 - Various Cleanup - Regions & If/Else of Script Path etc..
    # 1.4.0 - 2019-07-22 - Fix Issue with duplicate Web Application Output
    # 1.4.5 - 2019-08-20 - Add Export of Service Application Proxy Groups and also export the Proxy Group per Web App
    # 1.5.0 - 2020-08-31 - Prepare & Cleanup for GITHUB release
    ########################################################################################################################################
.SYNOPSIS
    ...
.DESCRIPTION
    This Script Exports All Essential Web Application Configuration Settings to XML, including Proxy group Association, Authentication Provider and Method, Alternate
    Access Mappings, Databases, Site Collections, IIS bindings and much more. All this information can be fed into another script to re-create the web applications
    on another farm, typically used for migration purposes.
.LINK
    ...
.EXAMPLE
    ...
#>
Param(
[parameter(mandatory=$false, HelpMessage = 'Farm Name will be used for the XML file Name')][string]$farmName = "PROD_2019" ,
[parameter(mandatory=$false, HelpMessage = 'Wildcard name of single Web App you want to exclude from XML export: e.g. "*MySite*"')][string]$webApplicationExclusion = "*OnlyTypeNameOfWebApplicationYouDontWantToExport*" ,
[parameter(mandatory=$false, HelpMessage = 'Filter out databases with name containing this string from XML export')][string]$dbNameExclFilter = "*OnlyTypeNameOfDatabasesYouDontWantToExport*" ,
[parameter(mandatory=$false, HelpMessage = 'Transcript this session?')][boolean]$EnableLogging = $false
)
########################################################################################################################################
#region - try
try
{
Clear-Host
Write-Host ""
Write-Host ("Script Started: ") -NoNewline -ForegroundColor Cyan
$startTime = (Get-Date)
Write-Host ((Get-Date -Format F)) -ForegroundColor DarkYellow
Write-Host ""
$ErrorActionPreference = "Stop";
###########################################################################################################################################
#region - HEADER (v1.2.7)
#region - HEADER VARIABLES
$InclusionScriptNames = @("")
$ScriptRequiresAdminPrivileges = $true
$ScriptRequiresSharePointAddin = $true
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
########################################################################################################################################
#region - Static Variable(s)
Write-Host "Reading Variables" -NoNewline
$xmlFileName = "WebApplicationSummary"
$encoding = "utf8"
$startDateTime = Get-date -Format yyyyMMddHHmmss
$scriptAction = "Export Web Application Settings and Bindings to XML"
$xmlFilePath = ($ScriptPath + "\" + $farmName + "_" + $xmlFileName + "_" + $startDateTime + ".xml")
$BlockedFileExtensionsDefaults = @("ade","adp","asa","ashx","asmx","asp","bas","bat","cdx","cer","chm","class","cmd","cnt","com","config","cpl","crt","csh","der","dll","exe","fxp","gadget","grp","hlp","hpj","hta","htr","htw","ida","idc","idq","ins","isp","its","jse","json","ksh","lnk","mad","maf","mag","mam","maq","mar","mas","mat","mau","mav","maw","mcf","mda","mdb","mde","mdt","mdw","mdz","msc","msh","msh1","msh1xml","msh2","msh2xml","mshxml","msi","ms-one-stub","msp","mst","ops","pcd","pif","pl","prf","prg","printer","ps1","ps1xml","ps2","ps2xml","psc1","psc2","pst","reg","rem","scf","scr","sct","shb","shs","shtm","shtml","soap","stm","svc","url","vb","vbe","vbs","vsix","ws","wsc","wsf","wsh","xamlx")
Write-Host "...Done" -ForegroundColor Green
#endregion - Static Variable(s)
########################################################################################################################################
#region - MAIN
#region - Output Script Title
Write-Host
Write-Host $scriptAction -ForegroundColor Cyan
for ($i = 0; $i -lt $scriptAction.Length; $i++)
    {
    Write-Host "=" -NoNewline -ForegroundColor Cyan
    }
Write-Host
#endregion - Output Script Title
#region - Load Additional Module(s)
Write-Host "Importing WebAdministration Module" -NoNewline
Import-Module WebAdministration #Import the IIS PowerShell module so we can interact with IIS       
Write-Host "...Done" -ForegroundColor Green
#endregion - Load Additional Module(s)
#region - Load Farm & Web Application Objects
Write-Host "Retrieving Farm Information" -NoNewline
$farm = Get-SPFarm -ErrorAction SilentlyContinue
$webApplications = Get-SPWebApplication -ErrorAction SilentlyContinue | Where-Object {($_.DisplayName -notin $webApplicationExclusion)}
Write-Host "...Done" -ForegroundColor Green
#endregion - Load Farm & Web Application Objects
#region - XML: Create Root Element
$xmlFile = New-Object Xml.XmlDocument
$xmlDeclaration = $xmlFile.CreateXmlDeclaration("1.0","utf-8",$null)
$xmlFile.InsertBefore($xmlDeclaration, $xmlFile.DocumentElement) | Out-Null
$null = $xmlRootElement
$xmlRootElement = $xmlFile.CreateElement("farm")
$xmlRootElement.SetAttribute("Name",$farm.Name)
$xmlRootElement.SetAttribute("BuildVersion",$farm.BuildVersion.ToString())
$xmlRootElement.SetAttribute("MajorVersion",$farm.BuildVersion.Major.ToString())
$xmlRootElement.SetAttribute("WebApplicationCount",$webApplications.Count)
$xmlRootElement.SetAttribute("NeedsUpgrade",$farm.NeedsUpgrade)
$xmlRootElement.SetAttribute("NeedsUpgradeIncludeChildren",$farm.NeedsUpgradeIncludeChildren)
$xmlFile.AppendChild($xmlRootElement) | Out-Null
$xmlFile.Save($xmlFilePath)
#endregion - XML: Create Root Element
#region - Iterate Through Web Application(s)
Write-Host ("Iterating through - " + $webApplications.Count + " - Web Application(s)")
$webAppCount = 1
foreach ($webApplication in $webApplications) 
    {
    Write-Host ("[" + $webAppCount + "/" + $webApplications.Count + "] `""+ $webApplication.DisplayName.ToUpper() + "`"") -ForegroundColor Cyan
    $webAppCount++
    $webApplicationContainsHnsc = ""
    Write-Host " + Getting Sites..." -NoNewline
    $webApplicationSites = Get-SPSite -WebApplication $webApplication.DisplayName -Limit All -WarningAction SilentlyContinue; Write-Host "..." -NoNewline
    $hnscUrls = ""; Write-Host "." -NoNewline
    #region - Check if web application contains hostnamed site collections
    foreach ($site in $webApplicationSites)
        {
        if ($site.HostHeaderIsSiteName)
            {
            $hnscUrls += $site.url; Write-Host "." -NoNewline
            }
        }
    if ($hnscUrls -eq "")
        {
        $webApplicationContainsHnsc = "false"; Write-Host "." -NoNewline
        }
    else
        {
        $webApplicationContainsHnsc = "true"; Write-Host "." -NoNewline
        }    
    Write-Host "...Done" -ForegroundColor Green
    #endregion - check if web application contains hostnamed site collections
    #region - XML: Create Web Application Element
    $xmlWebApplicationElement = $xmlFile.CreateElement("WebApplication")
    $xmlWebApplicationElement.SetAttribute("DisplayName",$webApplication.DisplayName)
    $xmlWebApplicationElement.SetAttribute("SitesCount",$webApplication.Sites.Count)
    $xmlWebApplicationElement.SetAttribute("MaxItemsPerThrottledOperation",$webApplication.MaxItemsPerThrottledOperation)
    $xmlWebApplicationElement.SetAttribute("MaximumFileSize",$webApplication.MaximumFileSize)
    $xmlWebApplicationElement.SetAttribute("MaxItemsPerThrottledOperationOverride",$webApplication.MaxItemsPerThrottledOperationOverride)
    $xmlWebApplicationElement.SetAttribute("MaxItemsPerThrottledOperationWarningLevel",$webApplication.MaxItemsPerThrottledOperationWarningLevel)
    $xmlWebApplicationElement.SetAttribute("MaxUniquePermScopesPerList",$webApplication.MaxUniquePermScopesPerList)
    $xmlWebApplicationElement.SetAttribute("DefaultTimeZone",$webApplication.DefaultTimeZone)
    $xmlWebApplicationElement.SetAttribute("AlertsEnabled",$webApplication.AlertsEnabled)
    $xmlWebApplicationElement.SetAttribute("AlertsMaximum",$webApplication.AlertsMaximum)
    $xmlWebApplicationElement.SetAttribute("AlertsMaximumQuerySet",$webApplication.AlertsMaximumQuerySet)
    $xmlWebApplicationElement.SetAttribute("RecycleBinEnabled",$webApplication.RecycleBinEnabled)
    $xmlWebApplicationElement.SetAttribute("RecycleBinCleanupEnabled",$webApplication.RecycleBinCleanupEnabled)
    $xmlWebApplicationElement.SetAttribute("RecycleBinRetentionPeriod",$webApplication.RecycleBinRetentionPeriod)
    $xmlWebApplicationElement.SetAttribute("SecondStageRecycleBinQuota",$webApplication.SecondStageRecycleBinQuota)
    $xmlWebApplicationElement.SetAttribute("AllowDesigner",$webApplication.AllowDesigner)
    $xmlWebApplicationElement.SetAttribute("AllowRevertFromTemplate",$webApplication.AllowRevertFromTemplate)
    $xmlWebApplicationElement.SetAttribute("AllowMasterPageEditing",$webApplication.AllowMasterPageEditing)
    $xmlWebApplicationElement.SetAttribute("ShowURLStructure",$webApplication.ShowURLStructure)
    $xmlWebApplicationElement.SetAttribute("BrowserFileHandling",$webApplication.BrowserFileHandling)
    $xmlWebApplicationElement.SetAttribute("ContainsHnsc",$webApplicationContainsHnsc)
    $xmlWebApplicationElement.SetAttribute("DefaultUrl",$webApplication.Url)
    $xmlWebApplicationElement.SetAttribute("ApplicationPoolName",$webApplication.ApplicationPool.Name)
    $xmlWebApplicationElement.SetAttribute("ApplicationPoolProcessAccountName",$webApplication.ApplicationPool.ProcessAccount.Name)
    $xmlWebApplicationElement.SetAttribute("MaxQueryLookupFields",$webApplication.MaxQueryLookupFields)
    $xmlWebApplicationElement.SetAttribute("Extensions",$webApplication.IisSettings.Keys.Count)
    $xmlWebApplicationElement.SetAttribute("ServiceApplicationProxyGroup",$webApplication.ServiceApplicationProxyGroup.FriendlyName)    
    if (!($BlockedFileExtensionDifferences = Compare-Object -ReferenceObject $BlockedFileExtensionsDefaults -DifferenceObject $webApplication.BlockedFileExtensions))
        {
        $xmlWebApplicationElement.SetAttribute("BlockedFileExtensions","default")
        }
    else
        {
        $xmlWebApplicationElement.SetAttribute("BlockedFileExtensions","changed")
        foreach ($difference in $BlockedFileExtensionDifferences)
            {
            $BlockedFileExtensionsRemoved = ""
            $BlockedFileExtensionsAdded = ""
            if ($difference.SideIndicator -eq "<=")
                {
                #Write-Host ("removed:" + $difference.InputObject)
                $BlockedFileExtensionsRemoved += ($difference.InputObject + ",")
                }
            else
                {
                #Write-Host ("added:" + $difference.InputObject)
                $BlockedFileExtensionsAdded += ($difference.InputObject + ",")
                }  
            $xmlWebApplicationElement.SetAttribute("BlockedFileExtensionsRemoved",$BlockedFileExtensionsRemoved.TrimEnd(","))
            $xmlWebApplicationElement.SetAttribute("BlockedFileExtensionsAdded",$BlockedFileExtensionsAdded.TrimEnd(","))      
            }
        }
    $xmlRootElement.AppendChild($xmlWebApplicationElement) | Out-Null
    $xmlFile.Save($xmlFilePath)
    Write-Host (" + Iterate through [" + $webApplication.IisSettings.Count + "] Extensions for this web application ") -ForegroundColor Gray
    #endregion - XML: Create Web Application Element
    #region - Iterate through Web Application extension(s) & Collect IIS Bindings
    foreach ($key in $webApplication.IisSettings.Keys) 
        {
        Write-Host (" - + `"" + $key + "`" Extension - Get Primary & Secondary AAM") -NoNewline
        $webApplicationAAMs = Get-SPAlternateURL -WebApplication $webApplication.DisplayName | Where-Object {$_.Zone -eq $key}
        $webApplicationPrimaryAAM = ""
        $webApplicationSecondaryAAMs = ""
        #region - XML: Create Extension Element
        $xmlWebApplicationExtensionElement = $xmlFile.CreateElement("extension")
        $xmlWebApplicationExtensionElement.SetAttribute("zone",$key)
        $xmlWebApplicationExtensionElement.SetAttribute("displayname",$webApplication.IisSettings.Item($key).ServerComment)
        foreach ($webApplicationAAM in $webApplicationAAMs)
            {
            $xmlWebApplicationExtensionAlternateUrlElement = $xmlFile.CreateElement("AlternateUrl")
            $xmlWebApplicationExtensionAlternateUrlElement.SetAttribute("Zone",$webApplicationAAM.Zone)
            $xmlWebApplicationExtensionAlternateUrlElement.SetAttribute("IncomingUrl",$webApplicationAAM.IncomingUrl)
            $xmlWebApplicationExtensionAlternateUrlElement.SetAttribute("PublicUrl",$webApplicationAAM.PublicUrl)
            $xmlWebApplicationExtensionElement.AppendChild($xmlWebApplicationExtensionAlternateUrlElement) | Out-Null
            $xmlFile.Save($xmlFilePath)
            if ($webApplicationAAM.IncomingUrl -eq $webApplicationAAM.PublicUrl)
                {
                $webApplicationPrimaryAAM = $webApplicationAAM.IncomingUrl
                }
            else
                {
                $webApplicationSecondaryAAMs += $webApplicationAAM.IncomingUrl
                }
            }
        Write-Host "...Done" -ForegroundColor Green
        #
        $xmlWebApplicationExtensionElement.SetAttribute("UseClaimsAuthentication",$webApplication.IisSettings.Item($key).UseClaimsAuthentication)
        $xmlWebApplicationExtensionElement.SetAttribute("UseWindowsIntegratedAuthentication",$webApplication.IisSettings.Item($key).UseWindowsIntegratedAuthentication)
        $xmlWebApplicationExtensionElement.SetAttribute("DisableKerberos",$webApplication.IisSettings.Item($key).DisableKerberos)
        $xmlWebApplicationExtensionElement.SetAttribute("UseBasicAuthentication",$webApplication.IisSettings.Item($key).UseBasicAuthentication)
        $xmlWebApplicationExtensionElement.SetAttribute("Anonymous",$webApplication.IisSettings.Item($key).AllowAnonymous)
        $xmlWebApplicationExtensionElement.SetAttribute("ServerComment",$webApplication.IisSettings.Item($key).ServerComment)
        $xmlWebApplicationElement.AppendChild($xmlWebApplicationExtensionElement) | Out-Null
        $xmlFile.Save($xmlFilePath)
        #endregion - XML: Create Extension Element
        #region - Get Extension Binding(s)
        $webSite = Get-Website | where {$_.Name -like $webApplication.IisSettings.Item($key).ServerComment}
        $siteBindings = $webSite.bindings.Collection
        $windowsAuthFilter = "/system.WebServer/security/authentication/windowsAuthentication"
        $useKernelMode = (Get-WebConfigurationProperty -pspath 'MACHINE/WEBROOT/APPHOST' -location $webSite.name -filter "$windowsAuthFilter" -name "useKernelMode").Value        
        #endregion - Get Extension Binding(s)
        #region - Iterate through Extension Binding(s)
        Write-Host (" - - + iterate through [" + $siteBindings.Count + "] binding(s) for this extension") -NoNewline -ForegroundColor Gray
        if ($siteBindings.Count -gt 0)
            {
            foreach ($siteBinding in $siteBindings) 
                {
                #region - Get Binding(s) Information
                $siteBindingString = $siteBinding.bindingInformation.ToString()
                $siteBindingPropertyArray = $siteBindingString.Split(":")
                $siteBindingIp = $siteBindingPropertyArray[0]
                $siteBindingPort = $siteBindingPropertyArray[1]
                $siteBindingHostHeader = $siteBindingPropertyArray[2]
                $siteBindingProtocol = $siteBinding.protocol
                $siteBindingUrl = ($siteBindingProtocol + "://" + $siteBindingHostHeader)
                $siteBindingIsPrimary = ""
                if ($siteBindingUrl -eq $webApplicationPrimaryAAM)
                    {
                    $siteBindingIsPrimary = "True"
                    $siteBindingIsAAM = "True"
                    }
                else
                    {
                    $siteBindingIsPrimary = "False"
                    if ($webApplicationSecondaryAAMs -like "*$siteBindingUrl*")
                        {
                        $siteBindingIsAAM = "True"
                        }
                    else
                        {
                        $siteBindingIsAAM = "False"
                        }
                    }
                if ($siteBindingHostHeader -ne "" -AND $hnscUrls -like "*$siteBindingHostHeader*" )
                    {
                    $siteBindingIsHNSC = "True"
                    }
                else 
                    {
                    $siteBindingIsHNSC = "False"
                    }
                #endregion - Get Binding(s) Information
                #region - XML: Create Extension Binding Element
                $xmlWebApplicationExtensionBindingElement = $xmlFile.CreateElement("binding")
                $xmlWebApplicationExtensionBindingElement.SetAttribute("ishnsc",$siteBindingIsHNSC)
                $xmlWebApplicationExtensionBindingElement.SetAttribute("useKernelMode",$useKernelMode)
                $xmlWebApplicationExtensionBindingElement.SetAttribute("aamexists",$siteBindingIsAAM)
                $xmlWebApplicationExtensionBindingElement.SetAttribute("primary",$siteBindingIsPrimary)
                $xmlWebApplicationExtensionBindingElement.SetAttribute("hostheader",$siteBindingHostHeader)
                $xmlWebApplicationExtensionBindingElement.SetAttribute("port",$siteBindingPort)
                $xmlWebApplicationExtensionBindingElement.SetAttribute("protocol",$siteBindingProtocol)
                $xmlWebApplicationExtensionBindingElement.SetAttribute("ip",$siteBindingIp)
                if ($siteBinding.protocol -eq "https")
                    {
                    $xmlWebApplicationExtensionBindingElement.SetAttribute("certificateHash",$siteBinding.certificateHash) # = ThumbPrint
                    $xmlWebApplicationExtensionBindingElement.SetAttribute("certificateStoreName",$siteBinding.certificateStoreName)
                    $xmlWebApplicationExtensionBindingElement.SetAttribute("sslFlags",$siteBinding.sslFlags) # = Require Server Name Indication Setting
                    }
                $xmlWebApplicationExtensionElement.AppendChild($xmlWebApplicationExtensionBindingElement) | Out-Null
                $xmlFile.Save($xmlFilePath)
                #endregion - XML: Create Extension Binding Element
                }
            }          
        Write-Host "...Done" -ForegroundColor Green        
        }  
        #endregion - Iterate through Extension Binding(s)
    #endregion - Iterate through Web Application extension(s) & Collect IIS Bindings       
    #region - Iterate through Content database(s)
    Write-Host " + Getting Database(s)" -NoNewLine
    $allDbs = $webApplication.ContentDatabases #Get-SPDatabase | Where-Object {($_.Type -eq "Content Database") -and ($_.WebApplication -notlike "*SPAdministrationWebApplication*")}
    $allFilteredDbs = $allDbs | Where-Object {$_.Name -notlike $dbNameExclFilter}    
    $webApplicationDbTotalSize = 0
    Write-Host "...Done" -ForegroundColor Green
    Write-Host (" - + Iterating through [" + $allDbs.Count + "] Content Databases") -ForegroundColor Gray
    $i = 0
    foreach ($db in $allFilteredDbs) 
        {
        if ($db.DisplayName.Length -gt $longestDbName) 
            {
            $longestDbName = $db.DisplayName.Length
            }
        }        
    foreach ($db in $allFilteredDbs) 
        {    
        $i++
        Write-Host (" - - + Content Database [" + $i + "/" + $allDbs.Count + "]: ") -NoNewline
        Write-Host ("`"" + $db.DisplayName + "`"") -ForegroundColor Yellow
        $dbSizeKB = $db.DiskSizeRequired
        $dbSizeMB = $dbSizeKB/1024/1024
        $dbSizeGB = $dbSizeKB/1024/1024/1024
        $dbSizeMB = [math]::Round($dbSizeMB,4)
        $dbSizeGB = [math]::Round($dbSizeGB,4)
        $webApplicationDbTotalSize += $dbSizeGB       
        $xmlWebApplicationDatabaseElement = $xmlFile.CreateElement("database")
        $xmlWebApplicationDatabaseElement.SetAttribute("DisplayName",$db.DisplayName)
        $xmlWebApplicationDatabaseElement.SetAttribute("SizeGB",$dbSizeGB)
        $xmlWebApplicationDatabaseElement.SetAttribute("Sites",$db.Sites.Count)
        $xmlWebApplicationDatabaseElement.SetAttribute("NeedsUpgrade",$db.NeedsUpgrade)
        $xmlWebApplicationDatabaseElement.SetAttribute("NeedsUpgradeIncludeChildren",$db.NeedsUpgradeIncludeChildren)
        $xmlWebApplicationDatabaseElement.SetAttribute("IsBackwardsCompatible",$db.IsBackwardsCompatible)
        $xmlWebApplicationDatabaseElement.SetAttribute("Server",$db.Server)
        $xmlWebApplicationDatabaseElement.SetAttribute("MaximumSiteCount",$db.MaximumSiteCount)
        $xmlWebApplicationDatabaseElement.SetAttribute("WarningSiteCount",$db.WarningSiteCount)
        $xmlWebApplicationDatabaseElement.SetAttribute("Status",$db.Status)
        $xmlWebApplicationDatabaseElement.SetAttribute("RemoteBlobStorageSettingsEnabled",$db.RemoteBlobStorageSettings.Enabled)
        $xmlWebApplicationDatabaseElement.SetAttribute("AvailabilityGroup",$db.AvailabilityGroup)
        $xmlWebApplicationDatabaseElement.SetAttribute("LastProfileSyncTime",$db.LastProfileSyncTime)
        $firstSiteCollectionUrl = $db.WebApplication.Url.Trim("/").ToString()
        if (($siteUrls = Get-SPSite -ContentDatabase $db.Name -Limit All -WarningAction SilentlyContinue) -ne $null)
			{			
			foreach ($siteUrl in $siteUrls)
                {                            
                if ($siteUrl.Url.ToString() -eq $firstSiteCollectionUrl)
                    {
                    $xmlWebApplicationDatabaseElement.SetAttribute("primaryDatabase",$true)
                    break
                    }
                else
                    {
                    $xmlWebApplicationDatabaseElement.SetAttribute("primaryDatabase",$false)
                    }                
                }
            }
        $xmlWebApplicationElement.AppendChild($xmlWebApplicationDatabaseElement) | Out-Null
        $xmlFile.Save($xmlFilePath) 
        Write-Host " - - - + Getting Site Collection(s)..." -NoNewline
        $spSites = $db | Get-SPSite -Limit all -WarningAction SilentlyContinue
        $j = 0
        Write-Host (" Iterating through [" + $spSites.count + "] SPSite(s):")
        #region - site collection
        foreach ($spSite in $spSites) 
            {        
            $j++
            $rootWeb = Get-SPWeb $spSite.Url
            Write-Host (" - - - - + [" + $j + "/" + $spSites.count +"] `"" + $rootWeb.Url +"`"" ) -NoNewline -ForegroundColor Gray
            $xmlWebApplicationDatabaseSiteCollectionElement = $xmlFile.CreateElement("sitecollection")
            $xmlWebApplicationDatabaseSiteCollectionElement.SetAttribute("Url",$spSite.Url)
            if ($rootWeb.WebTemplate)
                {
                $xmlWebApplicationDatabaseSiteCollectionElement.SetAttribute("Template",$rootWeb.WebTemplate)
                $xmlWebApplicationDatabaseSiteCollectionElement.SetAttribute("AllWebs",$spSite.AllWebs.Count.ToString())
                }
            else
                {
                $xmlWebApplicationDatabaseSiteCollectionElement.SetAttribute("Template","none")
                $xmlWebApplicationDatabaseSiteCollectionElement.SetAttribute("AllWebs",0)
                }
            $xmlWebApplicationDatabaseSiteCollectionElement.SetAttribute("CompatibilityLevel",$spSite.CompatibilityLevel)
            $xmlWebApplicationDatabaseSiteCollectionElement.SetAttribute("NeedsUpgrade",$spSite.NeedsUpgrade)
            $xmlWebApplicationDatabaseSiteCollectionElement.SetAttribute("NeedsUpgradeIncludeChildren",$spSite.NeedsUpgradeIncludeChildren)
            $xmlWebApplicationDatabaseSiteCollectionElement.SetAttribute("HostHeaderIsSiteName",$spSite.HostHeaderIsSiteName)
            $xmlWebApplicationDatabaseSiteCollectionElement.SetAttribute("SchemaVersion",$spSite.SchemaVersion.ToString())
            $xmlWebApplicationDatabaseSiteCollectionElement.SetAttribute("RecycleBinAllItems",$spSite.RecycleBin.Count)
            $xmlWebApplicationDatabaseSiteCollectionElement.SetAttribute("RecycleBinWebs",($spSite.RecycleBin | where {$_.ItemType -like "web"}).count)
                $spSiteSizeKB = $spSite.Usage.Storage
                $spSiteSizeMB = $spSiteSizeKB/1024/1024
                $spSiteSizeGB = $spSiteSizeKB/1024/1024/1024
                $spSiteSizeMB = [math]::Round($spSiteSizeMB,2)
                $spSiteSizeGB = [math]::Round($spSiteSizeGB,2)
            $xmlWebApplicationDatabaseSiteCollectionElement.SetAttribute("SizeGB",$spSiteSizeGB)
            $xmlWebApplicationDatabaseElement.AppendChild($xmlWebApplicationDatabaseSiteCollectionElement) | Out-Null
            $xmlFile.Save($xmlFilePath) 
            Write-Host "...Done" -ForegroundColor Green
        }
        #endregion - site collection
    }
    #endregion - Iterate through Content database(s)
$xmlWebApplicationElement.SetAttribute("SizeGB",$webApplicationDbTotalSize)  
$xmlFile.Save($xmlFilePath) 
}
Write-Host "Iteration through all *filtered* Web Applications - COMPLETE!" -ForegroundColor Green
Write-Host
#endregion - Iterate through Web Application(s)
#region - Iterate Through Service Application(s)
Write-Host "Service Applications:" -ForegroundColor Cyan
Write-Host "Retrieving Service Applications" -NoNewline
$serviceApplications = Get-SPServiceApplication -ErrorAction SilentlyContinue
Write-Host "...Done" -ForegroundColor Green
foreach ($serviceApplication in $serviceApplications)
    {
    Write-Host ("Processing: `"" + $serviceApplication.DisplayName + "`" (Type: `"" + $serviceApplication.TypeName + "`") ") -NoNewline
    $xmlServiceApplicationElement = $xmlFile.CreateElement("serviceapplication")
    $xmlServiceApplicationElement.SetAttribute("TypeName",$serviceApplication.TypeName)
    $xmlServiceApplicationElement.SetAttribute("DisplayName",$serviceApplication.DisplayName)
    $xmlServiceApplicationElement.SetAttribute("Id",$serviceApplication.Id)
    $xmlServiceApplicationElement.SetAttribute("ApplicationPoolName",$serviceApplication.ApplicationPool.Name)
    $xmlServiceApplicationElement.SetAttribute("ApplicationPoolAccount",$serviceApplication.ApplicationPool.ProcessAccount.Name)
    if ($serviceApplication.TypeName -eq "Search Service Application")
        {
        $xmlServiceApplicationElement.SetAttribute("IndexLocation",$serviceApplication.AdminComponent.IndexLocation)
        $xmlServiceApplicationElement.SetAttribute("ComponentCount",$serviceApplication.Topologies.ComponentCount)
        $xmlServiceApplicationElement.SetAttribute("CloudIndex",$serviceApplication.CloudIndex)
        }
    else
        {
        if ($serviceApplication.Database.Name)
            {
            $xmlServiceApplicationElement.SetAttribute("Database",$serviceApplication.Database.Name)
            }
        } 
    $xmlServiceApplicationElement.SetAttribute("Server",$db.Server)
    $xmlRootElement.AppendChild($xmlServiceApplicationElement) | Out-Null
    $xmlFile.Save($xmlFilePath) 
    Write-Host "...Done" -ForegroundColor Green
    }
Write-Host "Service Applications - Done" -ForegroundColor Green
#endregion - Iterate Through Service Application(s)
#region - Iterate Through Service Application(s) Proxy Group(s)
Write-Host "Retrieving Service Application Proxy Group(s):" -ForegroundColor Cyan
$serviceApplicationProxyGroups = Get-SPServiceApplicationProxyGroup
Write-Host "...Done" -ForegroundColor Green
foreach ($serviceApplicationProxyGroup in $serviceApplicationProxyGroups)
    {
    $xmlServiceApplicationProxyGroupElement = $xmlFile.CreateElement("serviceapplicationproxygroup")
    $xmlServiceApplicationProxyGroupElement.SetAttribute("FriendlyName",$SPServiceApplicationProxyGroup.FriendlyName)
    $xmlServiceApplicationProxyGroupElement.SetAttribute("Proxies",$SPServiceApplicationProxyGroup.Proxies)
    $xmlRootElement.AppendChild($xmlServiceApplicationProxyGroupElement) | Out-Null
    $xmlFile.Save($xmlFilePath) 
    }
Write-Host "Retrieving Service Application Proxy Group(s) - Done" -ForegroundColor Green
Write-Host ("The XML Output can be found here:" + $xmlFilePath) -ForegroundColor Yellow
#endregion - Iterate Through Service Application(s) Proxy Group(s)
#endregion - MAIN
########################################################################################################################################
}
#endregion - try
########################################################################################################################################
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
    $endTime = (Get-Date)
    if ($StartSPAssignMentVariable) {Stop-SPAssignment -Global | Out-Null}    
    $ErrorActionPreference = "Continue";
    Write-Host ""    
    Write-Host ("Script Ended: ") -NoNewline -ForegroundColor DarkCyan
    Write-Host ((Get-Date -Format F)) -ForegroundColor DarkYellow -NoNewline      
    Write-Host (" (and took: ") -ForegroundColor Gray -NoNewline
    $timeSpan = ($($endTime - $startTime).TotalSeconds)
    Write-Host ([math]::Round($timeSpan,3)) -NoNewline
    Write-Host (" Seconds)") -ForegroundColor Gray
    if ($EnableLogging) {Stop-Transcript}
    #endregion - Finally
}
#endregion - Catch & Finally
#######################################################################################################################################