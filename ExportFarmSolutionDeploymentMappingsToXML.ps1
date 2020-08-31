<#
.NOTES
    ########################################################################################################################################
    # Author: Jim B.
    ########################################################################################################################################
    # Revision(s)
    # 0.0.1 - 2018-10-17 - Initial Version
    # 1.0.0 - 2020-08-31 - Initial Commit to GITHUB
    ########################################################################################################################################
.SYNOPSIS
    ...
.DESCRIPTION
    This script exports Farm Solution Deployment mappings to XML.
.LINK
    ...
.EXAMPLE
    ...
.EXAMPLE
    ...
#>
Param(
[parameter(mandatory=$false, HelpMessage = 'String will be used for the Export Folder Name')][string]$xmlOutputFileName = "FarmSolutionDeploymentMappings.xml" ,
[parameter(mandatory=$false, HelpMessage = 'Example: "*nintex*"')][string]$SolutionsExportExclusionFilterWildcard1 ,
[parameter(mandatory=$false, HelpMessage = 'Example: "*bamboo*"')][string]$SolutionsExportExclusionFilterWildcard2 ,
[parameter(mandatory=$false, HelpMessage = 'Transcript this session?')][boolean]$EnableLogging = $false
)
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
########################################################################################################################################
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
#region - LOCAL VARIABLES
$scriptAction = "Export Farm Solution Deployment Mappings to XML"
#
$xmlFilePath = ($ScriptPath + "\" + $xmlOutputFileName + "_" + (Get-date -Format yyyyMMddHHmmss) + ".xml")
#
#
#$farmName = (Get-SPFarm).Name.Substring(0, ((Get-SPFarm).Name.LastIndexOf("_")))
#$xmlFilePath = ($ScriptPath + "\" + $farmName + "_" + $xmlOutputFileName + "_" + $startDateTime + ".xml")
#
#
#endregion - LOCAL  VARIABLES
########################################################################################################################################
#region - Function(s)
#endregion - Function(s)
########################################################################################################################################
#region - MAIN
#region - Output Script Title
Write-Host
Write-Host $scriptAction -ForegroundColor Cyan
for ($i = 0; $i -le $scriptAction.Length; $i++)
    {
    Write-Host "=" -NoNewline -ForegroundColor Cyan
    }
Write-Host
#endregion - Output Script Title
#region - XML: Create Root Element
$farm = Get-SPFarm
$xmlFile = New-Object Xml.XmlDocument
$xmlDeclaration = $xmlFile.CreateXmlDeclaration("1.0","utf-8",$null)
$xmlFile.InsertBefore($xmlDeclaration, $xmlFile.DocumentElement) | Out-Null
$null = $xmlRootElement
$xmlRootElement = $xmlFile.CreateElement("Farm")
$xmlRootElement.SetAttribute("Name",$farm.Name)
$xmlRootElement.SetAttribute("BuildVersion",$farm.BuildVersion.ToString())
$xmlRootElement.SetAttribute("MajorVersion",$farm.BuildVersion.Major.ToString())
$xmlFile.AppendChild($xmlRootElement) | Out-Null
$xmlFile.Save($xmlFilePath)
#endregion - XML: Create Root Element
#region - Get-SPSolution
$spSolutions = Get-SPSolution | where {($_.Name -notlike $SolutionsExportExclusionFilterWildcard1) -and ($_.Name -notlike $SolutionsExportExclusionFilterWildcard2)}
if ($spSolutions -eq $null)
    {
    Write-Host ("No Solutions Found - Stopping Script!") -ForegroundColor Yellow
    exit
    }
[int]$longestSolutionFileName = 0
foreach ($spSolution in $spSolutions)    
        {
        if ($spSolution.Name.Length -gt $longestSolutionFileName) 
            {
            $longestSolutionFileName = $spSolution.Name.Length
            }
        }
#endregion - Get-SPSolution
#region - XML: Create Solutions Element
$xmlSolutionsElement = $xmlFile.CreateElement("Solutions")
$xmlSolutionsElement.SetAttribute("Count",$spSolutions.Count.ToString())
$xmlRootElement.AppendChild($xmlSolutionsElement) | Out-Null
$xmlFile.Save($xmlFilePath)
#endregion - XML: Create Solutions Element
#region - Iterate through Solution(s)
foreach ($spSolution in $spSolutions)
    {
    #region - XML: Create Solution Element
    $xmlSolutionElement = $xmlFile.CreateElement("Solution")
    $xmlSolutionElement.SetAttribute("DisplayName",$spSolution.DisplayName)
    $xmlSolutionElement.SetAttribute("Id",$spSolution.ID.Guid.ToString())
    $xmlSolutionElement.SetAttribute("Deployed",$spSolution.Deployed)
    $xmlSolutionsElement.AppendChild($xmlSolutionElement) | Out-Null
    $xmlFile.Save($xmlFilePath)
    #endregion - XML: Create Solution Element
    if (!($spSolution.Deployed))
        {
        $xmlSolutionElement.SetAttribute("Global",$false)
        }
    else
        {
        if (!($spSolution.DeployedWebApplications))
            {                            
            $xmlSolutionElement.SetAttribute("Global",$true)
            }
        else
            {
            $xmlSolutionElement.SetAttribute("Global",$false)
            #
            $xmlDeployedElement = $xmlFile.CreateElement("Deployed")
            $xmlSolutionElement.AppendChild($xmlDeployedElement) | Out-Null
            $xmlFile.Save($xmlFilePath)
            foreach ($deployedWebApp in $spSolution.DeployedWebApplications)
                {                                    
                $xmlDeployedWebAppElement = $xmlFile.CreateElement("WebApp")
                $xmlDeployedElement.AppendChild($xmlDeployedWebAppElement) | Out-Null
                $xmlDeployedElement.SetAttribute("Url",$deployedWebApp.Url)
                $xmlFile.Save($xmlFilePath)
                }            
            }            
        }        
    foreach ($featureGroups in Get-SPFeature | Where-Object {$_.SolutionID -eq $spSolution.id} | Group-Object SolutionId)
        {
        [int]$spacesSolutionFileName = (($longestSolutionFileName + 2) - $spSolution.Name.Length)
        Write-Host ("[") -NoNewline
        Write-Host ("{0}{1,$spacesSolutionFileName}" -f $spSolution.Name, ":") -NoNewline
        Write-Host (" - " + $spSolution.ID + " Count:" + $featureGroups.Count + "]")
        [int]$longestFeatureNameLength = 0    
        #region - XML: Create Solutions Element
        $xmlFeaturesElement = $xmlFile.CreateElement("Features")
        $xmlFeaturesElement.SetAttribute("Count",$featureGroups.Count)
        $xmlSolutionElement.AppendChild($xmlFeaturesElement) | Out-Null
        $xmlFile.Save($xmlFilePath)
        #endregion - XML: Create Solutions Element                    
        [int]$longestFeatureNameLength = ($featureGroups.Group | Sort-Object -Property DisplayName -Descending | Select-Object -First 1).DisplayName.Length
        foreach ($fd in $featureGroups.Group | sort DisplayName ) 
            {            
            [int]$spacedFeatureName = (($longestFeatureNameLength + 3) - $fd.DisplayName.Length)
            Write-Host (" - ") -NoNewline
            Write-Host ("{0}{1,$spacedFeatureName}" -f $fd.DisplayName, ": ") -NoNewline
            Write-Host $fd.DisplayName '-' $fd.Id '(' $fd.Scope ')'
            #region - XML: Create Feature Element
            $xmlFeatureElement = $xmlFile.CreateElement("Feature")
            $xmlFeatureElement.SetAttribute("DisplayName", $fd.DisplayName)
            $xmlFeatureElement.SetAttribute("Id", $fd.Id)
            $xmlFeatureElement.SetAttribute("Scope", $fd.Scope)
            $xmlFeatureElement.SetAttribute("CompatibilityLevel", $fd.CompatibilityLevel)            
            $xmlFeaturesElement.AppendChild($xmlFeaturesElement) | Out-Null
            $xmlFile.Save($xmlFilePath)
            #endregion - XML: Create Feature Element
            }
        }
    }
Write-Host
Write-Host ("The XML Output can be found here:" + $xmlFilePath) -ForegroundColor Yellow
#endregion - Iterate through Solutions
#endregion - MAIN Scripting Body
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