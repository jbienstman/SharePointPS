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
    This script exports Farm Solution WSP files to a new folder created in the Script Location.
.LINK
    ...
.EXAMPLE
    ...
.EXAMPLE
    ...
#>
Param(
[parameter(mandatory=$false, HelpMessage = 'String will be used for the Export Folder Name')][string]$farmName = "PROD_2019" ,
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
$scriptAction = "Export Farm Solution WSP Files:"
$ExportSolutionsPath = ($ScriptPath + "\" + $farmName + "_WSP")
#endregion - LOCAL  VARIABLES
########################################################################################################################################
#region - Function(s)
Function AskYesNoQuestion {
<#
.EXAMPLE
AskYesNoQuestion ("Your Question Text Here?")
#>
    Param (
        [Parameter(Mandatory=$true)][string]$Question
    )
    $Choice1 = "y"
    $Choice2 = "n"
    $QuestionSuffix = "[$Choice1/$Choice2]"
    Do {[string]$CheckAnswer = Read-Host ($Question + $QuestionSuffix)}
    Until ($CheckAnswer -eq $Choice1 -or $CheckAnswer -eq $Choice2)
    Switch ($CheckAnswer)
        {            
            y {Return $True}
            n {Return $False}
        }  
}
#endregion - Function(s)
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
#region - Get-SPSolution
Write-Host ("To Location: " + $ExportSolutionsPath) -ForegroundColor White
$Solutions = Get-SPSolution | where {($_.Name -notlike $SolutionsExportExclusionFilterWildcard1) -and ($_.Name -notlike $SolutionsExportExclusionFilterWildcard2)}
#endregion - Get-SPSolution
#region - Export Solution WSP Files (if any)
if ($Solutions -ne $null)
    {
    #region - Create Output Path if not exists
    if (!(Test-Path $ExportSolutionsPath))
        {
        $createSubFolder = AskYesNoQuestion ("Do you want to create a the folder: " + $ExportSolutionsPath)
        if ($createSubFolder -eq $true)
            {
            New-Item -Path $ExportSolutionsPath -ItemType Directory | Out-Null
            }
        }
    #endregion - Create Output Path if not exists
    #region - Iterate through Solutions & Export WSP File
    foreach ($Solution in $Solutions)
        {
        $SolFile = $Solution.SolutionFile
        $SolName = $Solution.Name
        $ExportSolutionsLoc = $ExportSolutionsPath +"\"+ $SolName 
        Write-Host (" - " + $SolName) -NoNewline
        $SolFile.SaveAs($ExportSolutionsLoc)
        Write-Host " ...Done" -ForegroundColor Green
        }
    #endregion - Iterate through Solutions & Export WSP File
    #region - Ouput Message
    Write-Host (" - Export COMPLETE! You can find the solutions here: ") -NoNewline -ForegroundColor White
    Write-Host ("`"" + $ExportSolutionsPath + "`"") -ForegroundColor Yellow
    Write-Host "`n"
    #endregion - Ouput Message
    }
else
    {
    Write-Host (" - No Solutions found to Export!") -ForegroundColor Yellow
    Write-Host "`n"
    }
#endregion - Export Solution WSP Files (if any)
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