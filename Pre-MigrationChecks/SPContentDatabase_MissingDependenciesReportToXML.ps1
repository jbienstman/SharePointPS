<#
.NOTES
    ########################################################################################################################################
    # Author: Jim B.
    ########################################################################################################################################
    # Revision(s)
    # 1.0.0 - 2018-10-09 - Initial Version
    # 1.2.0 - 2019-07-19 - Updated XML Constructor
    # 1.3.0 - 2019-07-31 - Cleanup unneeded variables and functions, added latest header section
    ########################################################################################################################################
.SYNOPSIS
    Export the Results of Test-SPContentDatabase for each content database in the farm.
.DESCRIPTION
    Export the Results of Test-SPContentDatabase for each content database in the farm.
.LINK
    ...
.EXAMPLE
    ...
.EXAMPLE
    ...
#>
########################################################################################################################################
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
###########################################################################################################################################
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
$scriptAction = "Missing Dependencies - Report to XML"
$xmlFilePathNamePrefix = "MissingDependenciesReport"
$startDateTime = Get-date -Format yyyyMMddHHmmss
$xmlFilePath = ($ScriptPath + "\" + $xmlFilePathNamePrefix + "_" + $startDateTime + ".xml")
#
#endregion - MAIN SCRIPT VARIABLES
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
#region - Get Top Level Stuff
$farm = Get-SPFarm
$databases = Get-SPDatabase | Where-Object {($_.Type -eq "Content Database") -and ($_.WebApplication -notlike "*SPAdministrationWebApplication*")}
#endregion - Get Top Level Stuff
#region - XML Farm
$null = $xmlRootElement
$xmlFile = New-Object Xml.XmlDocument
$xmlDeclaration = $xmlFile.CreateXmlDeclaration("1.0","utf-8",$null)
$xmlFile.InsertBefore($xmlDeclaration, $xmlFile.DocumentElement) | Out-Null
$xmlRootElement = $xmlFile.CreateElement("farm")
$xmlRootElement.SetAttribute("Name",$farm.Name)
$xmlRootElement.SetAttribute("BuildVersion",$farm.BuildVersion.ToString())
$xmlRootElement.SetAttribute("MajorVersion",$farm.BuildVersion.Major.ToString())
$xmlRootElement.SetAttribute("Databases",$databases.Count)
$xmlFile.AppendChild($xmlRootElement) | Out-Null
$xmlFile.Save($xmlFilePath)
#endregion - XML Farm
#region - Content Database Iteration
foreach ($database in $databases)
    {
    $databaseErrors = Test-SPContentDatabase -Name $database.DisplayName -WebApplication $database.WebApplication.Url -ServerInstance $database.Server
    $null = $xmlDatabaseElement
    $xmlDatabaseElement = $xmlFile.CreateElement("Database")
    $xmlDatabaseElement.SetAttribute("Name",$database.Name)
    $xmlDatabaseElement.SetAttribute("ErrorCount",$databaseErrors.Count)
    $xmlRootElement.AppendChild($xmlDatabaseElement) | Out-Null
    $xmlFile.Save($xmlFilePath)
    Write-Host (" - " + $database.DisplayName + " - ") -NoNewline
    Write-Host ( $databaseErrors.Count) -ForegroundColor Yellow
    If ($databaseErrors.Count -gt 0)
        {
        foreach ($databaseError in $databaseErrors)
            {
            if ($databaseError.Remedy) 
                {
                $Remedy = $databaseError.Remedy.Replace("'","&apos;")
                }
            else
                {
                $Remedy = ""
                }

            $null = $xmlMissingDependencyElement
            $xmlMissingDependencyElement = $xmlFile.CreateElement("MissingDependency")
            $xmlMissingDependencyElement.SetAttribute("Category",$databaseError.Category)
            $xmlMissingDependencyElement.SetAttribute("Error",$databaseError.Error)
            $xmlMissingDependencyElement.SetAttribute("UpgradeBlocking",$databaseError.UpgradeBlocking)
            $xmlMissingDependencyElement.SetAttribute("Message",$databaseError.Message.Replace("'","&apos;"))
            $xmlMissingDependencyElement.SetAttribute("Remedy",$Remedy)
            $xmlMissingDependencyElement.SetAttribute("Locations",$databaseError.Locations)
            $xmlDatabaseElement.AppendChild($xmlMissingDependencyElement)
            $xmlFile.Save($xmlFilePath)
            }
        }
    }
#endregion - Content Database Iteration
Write-Host ("Output can be found here: `"" + $xmlFilePath + "`"") -ForegroundColor Yellow
#endregion - MAIN
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