<#
.NOTES
    #######################################################################################################################################
    # Author: jimmy.bienstman@microsoft.com
    #######################################################################################################################################
    # Revision(s)
    # 0.0.1 - 2018-10-09 - Initial Version
    # 1.0.0 - 2018-10-16 - Many Changes - Cleanup, Stability, Additional Error Handling, 
                           REPORT Output to CSV using Objects, no longer need to feed the XML file location - Menu shows the latest one on top
                           and much more...
    # 1.0.2 - 2018-10-16 - Added untested, partial handling of the 8 different types of Missing Assemblies
    # 1.0.6 - 2018-10-18 - Cleanup Code, working cleanup of Missing Features and Setup Files
    # 1.0.7 - 2018-11-12 - Replaced Header 1.2.4 with 1.2.5, Tested on additional environments - improving stability
    # 1.1.0 - 2019-07-31 - Cleanup, Updated Header version 1.2.7, only included needed functions
    #######################################################################################################################################
.SYNOPSIS
    ...
.DESCRIPTION
    This Script parses the XML output of the "SPContentDatabase_MissingDependenciesReportToXML.ps1" script and reports more detailed information
    such as the "Url" and "Scope" of a Feature, Setup File or WebPart.
    Cleaning Options:
    - deleteMissingFeatures: $true - (Will still prompt you for approval)
    - deleteMissingSetupFiles: $true - (Will still prompt you for approval)
.LINK
    ...
.EXAMPLE
    Just place the XML in the same folder as this script on one of the affected WFE servers in the farm
.EXAMPLE
    ...
#>
###########################################################################################################################################
#region - TRY
try 
{
###########################################################################################################################################
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
#region - Function(s)
Function MenuListChoice {
    <#
    .EXAMPLE
        $menuChoiceValue = MenuListChoice -menuListOptions $arrayOfChoices -Title 'Menu Title' -Message ("Please type a number from the list and press [ENTER]") -checkParameterName "Name"
    #>
    Param (
        [Parameter(Mandatory=$true)][array]$menuListOptions ,
        [Parameter(Mandatory=$false)][string]$Title = "Make your choice from the list below:" ,
        [Parameter(Mandatory=$false)][string]$checkParameterName = "Name" ,
        [Parameter(Mandatory=$false)][string]$Message = "Please type a number from the list and press [ENTER]",
        [Parameter(Mandatory=$false)][string]$TitleForegroundColor = "Cyan" ,
        [Parameter(Mandatory=$false)][string]$ChoiceForegroundColor = "Yellow"
    )   
        #region - Initialize Menu Array & Write Title
        $menu = @{}        
        Write-Host
        Write-Host ($Title) -ForegroundColor $TitleForegroundColor
        for ($i = 0; $i -lt $Title.Length; $i++)
        {
        Write-Host "-" -NoNewline -ForegroundColor $TitleForegroundColor;
        }
        Write-Host        
        #endregion - Initialize Menu Array & Write Title
        #region - Build Menu
        $digits = $menuListOptions.Count.ToString().Length
        for ($i=1 ; $i -le $menuListOptions.Count ; $i++) 
            {
            if ($checkParameterName = "None")
                {
                Write-Host ("{0:d$digits}" -f $i) -NoNewline -ForegroundColor $ChoiceForegroundColor
                Write-Host (" - " +  $($menuListOptions[$i-1]))
                $menu.Add($i,($menuListOptions[$i-1]))
                }
            else
                {
                Write-Host ("{0:d$digits}" -f $i) -NoNewline -ForegroundColor $ChoiceForegroundColor
                Write-Host (" - " +  $($menuListOptions[$i-1].$checkParameterName))
                $menu.Add($i,($menuListOptions[$i-1].$checkParameterName))
                }
            }
        Write-Host ""
        #endregion - Build Menu
        #region - Get & Validate Selection
        do {
            try {
                $validChoice = $true        
                [int]$choice = Read-Host -Prompt $Message
                }
            catch 
                {
                $validChoice = $false
                }
            }
        until (($choice -ge 0 -and $choice -le $menu.Count) -and $validChoice)
        $selection = $menu.Item($choice)
        #endregion - Get & Validate Selection
        #region - Return Value
        Return $selection
        #endregion - Return Value
}
Function WriteCharLine { 
    <#
    .EXAMPLE
        WriteCharLine -lineFillCharacter "-" -lineLength 80 -lineForegroundColor Gray
    #>
    Param (
        [Parameter(Mandatory=$true)][string]$lineFillCharacter ,
        [Parameter(Mandatory=$true)][int]$lineLength ,
        [Parameter(Mandatory=$false)][string]$lineColor = "White"
    )
    for ($i = 0; $i -lt $lineLength; $i++)
        {
        Write-Host $lineFillCharacter -NoNewline -ForegroundColor $lineColor;
        }        
    Write-Host;
}
Function Run-SQLQuery {
    <#
    .EXAMPLE
        Run-SQLQuery -SqlServer "SQLSERVERNAME01" -SqlDatabase "DATABASENAME01" -SqlQuery "SELECT column FROM table WITH (NoLOCK) WHERE column IN ('matchterm')"
    #>
    Param (
        [Parameter(Mandatory=$true)][string]$SqlServer ,
        [Parameter(Mandatory=$true)][string]$SqlDatabase ,
        [Parameter(Mandatory=$true)][string]$SqlQuery
    )
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server =" + $SqlServer + "; Database =" + $SqlDatabase + "; Integrated Security = True"
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.CommandText = $SqlQuery
    $SqlCmd.Connection = $SqlConnection
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd
    $DataSet = New-Object System.Data.DataSet
    $SqlAdapter.Fill($DataSet)
    $SqlConnection.Close()
    $DataSet.Tables[0]
    #help Run-SQLQuery -Examples
    }
Function AskYesNoQuestion {
    <#
    .EXAMPLE
        AskYesNoQuestion ("Your Question Text Here?")
    #>
    Param (
        [Parameter(Mandatory=$true)][string]$Question ,
        [Parameter(Mandatory=$false)][string]$ForegroundColor = "White",        
        [Parameter(Mandatory=$false)][string]$Choice1 = "y" ,
        [Parameter(Mandatory=$false)][string]$Choice2 = "n"
    )    
    $QuestionSuffix = "[$Choice1/$Choice2]"
    Do {Write-Host ($Question) -ForegroundColor $ForegroundColor -NoNewline;[string]$CheckAnswer = Read-Host $QuestionSuffix}
    Until ($CheckAnswer -eq $Choice1 -or $CheckAnswer -eq $Choice2)
    Switch ($CheckAnswer)
        {            
            $Choice1 {Return $True}
            $Choice2 {Return $False}
        }  
}
#endregion - Function(s)
###########################################################################################################################################
#region - MAIN SCRIPT VARIABLES
$scriptAction = "Missing Dependencies Parser"
#
$startDateTime = Get-date -Format yyyyMMddHHmmss
$missingDependenciesXMLReportFilter = "MissingDependenciesReport*.xml"
#
$deleteMissingFeatures = $False
$deleteMissingSetupFiles = $False
#
#endregion - MAIN SCRIPT VARIABLES
###########################################################################################################################################
#region - MAIN
#region - Output Script Title
Write-Host
if ($deleteMissingFeatures -ne $True -and $deleteMissingSetupFiles -ne $True)
    {
    $scriptAction = ($scriptAction + " (REPORT ONLY MODE)")
    }
else
    {
    $scriptAction = ($scriptAction + " (CLEANUP MODE)")
    }
Write-Host
Write-Host $scriptAction -ForegroundColor Cyan
for ($i = 0; $i -lt $scriptAction.Length; $i++)
    {
    Write-Host "=" -NoNewline -ForegroundColor Cyan
    }
Write-Host
#endregion - Output Script Title
#region - Retrieve Existing MissingDependenciesReport*.XML Files
$sortProperty = "CreationTime"
$MissingDependenciesReportsInCurrentLocation = Get-ChildItem -Path $ScriptPath -Filter $missingDependenciesXMLReportFilter | Sort-Object -Property $sortProperty -Descending
if ($null -ne $MissingDependenciesReportsInCurrentLocation)
    {
    $xmlOutputFileName = MenuListChoice -menuListOptions $MissingDependenciesReportsInCurrentLocation -Title "Reports with newest `"$sortProperty`" are on Top of the list!" -Message "Type number and press [ENTER] to continue."
    Write-Host ("Parsing XML: `"" + $xmlOutputFileName + "`"") -NoNewline -ForegroundColor Gray
    [xml]$xmlFile = Get-Content ($ScriptPath + "\" + $xmlOutputFileName)
    Write-Host (" ...Done") -ForegroundColor DarkGreen
    $errorDatabases = $xmlFile.Farm.Database | Where-Object {$_.ErrorCount -ne 0}
    }
else
    {
    Write-Host (" - WARNING: Cannot find any file matching filter `"$missingDependenciesXMLReportFilter`" in: `"" + $ScriptPath + "`"") -NoNewline -ForegroundColor Yellow
    Write-Host
    exit
    }
#endregion - Retrieve Existing MissingDependenciesReport*.XML Files
#region - Check for Errors
if ($null -eq $errorDatabases)
    {
    Write-Host (" - " + $xmlOutputFileName + ": Has no errors for any of the Content Databases tested - Good Job!") -ForegroundColor Green
    exit
    }
#endregion - Check for Errors
$missingDependenciesObjectArray = @()
foreach ($errorDatabase in $errorDatabases)
    {
    #region - Draw a Line
    Write-Host  
    $lineFillCharacter = "-"
    $lineLength = 80
    $lineColor = "Gray"
    WriteCharLine -lineFillCharacter $lineFillCharacter -lineLength $lineLength -lineColor $lineColor
    #endregion - Draw a Line
    Write-Host ("DATABASE `"" ) -ForegroundColor White -NoNewline
    Write-Host ($errorDatabase.Name.ToUpper()) -ForegroundColor Cyan -NoNewline
    Write-Host ("`"") -ForegroundColor White
    #region - Missing Dependency Iteration
    foreach($missingDependency in $errorDatabase.MissingDependency)
        {
        #region - MISSING FEATURE
        if ($missingDependency.Category -eq "MissingFeature")
            {
            #region - REGEX
            $startMatches = ([regex]'\[').Matches($missingDependency.Message);
            $endMatches = ([regex]'\]').Matches($missingDependency.Message);
            $StartIndex0 = $startMatches[0].Index + 1
            $EndIndex0 = $endMatches[0].Index - $StartIndex0
            $StartIndex1 = $startMatches[1].Index + 1
            $EndIndex1 = $endMatches[1].Index - $StartIndex1
            #endregion - REGEX
            #region - Get Missing Dependency Info
            $databaseName = $missingDependency.Message.Substring($StartIndex0,$EndIndex0)            
            $featureDesignator = $missingDependency.Message.Substring($StartIndex1 -6 ,$StartIndex1 - ($StartIndex1 - 2))   
            if ($featureDesignator -eq "id")
                {
                $featureId = $missingDependency.Message.Substring($StartIndex1,$EndIndex1)
                }
            else
                {
                $StartIndex2 = $startMatches[2].Index + 1
                $EndIndex2 = $endMatches[2].Index - $StartIndex2
                $featureId = $missingDependency.Message.Substring($StartIndex2,$EndIndex2)
                }
            #endregion - Get Missing Dependency Info
            #region - SQL Query            
            $databaseServer = (Get-SPContentDatabase $errorDatabase.Name).Server
            $sqlQueryFeatures = "SELECT DISTINCT F.SiteId, F.WebId, F.FeatureId FROM [dbo].[Features] AS F WITH (NoLOCK) WHERE F.FeatureId IN ('$featureId')"            
            $missingFeatures = Run-SQLQuery -SqlServer $databaseServer -SqlDatabase $errorDatabase.Name -SqlQuery $sqlQueryFeatures
            #endregion - SQL Query
            #region - Iteration
            Write-Host ("+ " + $missingDependency.Category + ":[" + $featureId + "]") -ForegroundColor Cyan -NoNewline
            if ($missingFeatures -ne 0) #When the Features would not be found in the DB Query - SQL would return ZERO "0"
                {    
                Write-Host (" - Found!") -ForegroundColor DarkCyan
                foreach ($missingFeature in $missingFeatures)            
                    {
                    if ($null -ne $missingFeature.FeatureId) #First Entry is typically $null - so we're skipping this one
                        { 
                        #region - Get Feature Title
                        $sqlQueryFeatureTitle = "SELECT FeatureTitle FROM [dbo].[FeatureTracking] WITH (NoLOCK) WHERE FeatureId IN ('$featureId')"
                        $featureTitleResult = Run-SQLQuery -SqlServer $databaseServer -SqlDatabase $errorDatabase.Name -SqlQuery $sqlQueryFeatureTitle                            
                        if($featureTitleResult -eq 0)
                            {
                            $featureTitle = "None"
                            }
                        else
                            {
                            $featureTitle = $featureTitleResult.FeatureTitle
                            }
                        #endregion - Get Feature Title
                        #region - Get Feature Url
                        if ($missingFeature.WebId.Guid -eq "00000000-0000-0000-0000-000000000000")
                            {
                            #Write-Host ($missingFeature.WebId.Guid + " - ") -NoNewline
                            $site = Get-SPSite -Identity $missingFeature.SiteId.ToString()     
                            $web = $site.RootWeb
                            $missingFeatureScope = "SPSite"
                            Write-Host (" - [$missingFeatureScope Scoped] [Title: `"" + $featureTitle + "`"] - URL: ") -NoNewline -ForegroundColor Gray
                            $missingFeatureUrl = $web.Url
                            Write-Host $missingFeatureUrl
                            }
                        else
                            {                            
                            $site = Get-SPSite -Identity $missingFeature.SiteId.ToString()                             
                            $web = Get-SPWeb -Site $missingFeature.SiteId.ToString() -Identity $missingFeature.WebId.ToString() -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                            $missingFeatureScope = "SPWeb"
                            if ($null -eq $web)
                                {
                                # = $True
                                $web = $site.RootWeb
                                Write-Host (" - [$missingFeatureScope Scoped] [Title: `"" + $featureTitle + "`"] - URL: ") -NoNewline -ForegroundColor Gray
                                $missingFeatureUrl = ($web.Url + " - Orphaned WebId: " + $missingFeature.WebId.ToString())                                  
                                Write-Host  $missingFeatureUrl -ForegroundColor Yellow
                                }
                            else
                                {
                                $missingFeatureUrl = $web.Url
                                Write-Host (" - [$missingFeatureScope Scoped] [Title: `"" + $featureTitle + "`"] - URL: ") -NoNewline -ForegroundColor Gray
                                Write-Host $missingFeatureUrl
                                }              
                            }                        
                        #endregion - Get Feature Url    
                        #region - Build Web Owner String
                        $webOwnerList = ""
                        foreach ($webOwner in $web.AssociatedOwnerGroup.Users)
                            {
                            $webOwnerList += ("`"" + $webOwner.DisplayName + "`",")
                            }    
                        #endregion - Build Web Owner String
                        #region - Build & Add Object to the Array       
                        $missingDependencyObject = New-Object –TypeName PSObject
                        $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Database -Value $errorDatabase.Name.ToUpper()                       
                        $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Category -Value $missingDependency.Category
                        $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Id -Value $featureId
                        $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Name -Value $featureTitle
                        $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Scope -Value $missingFeatureScope
                        $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Url -Value $web.Url
                        $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Occurences -Value 1
                        $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Owners -Value $webOwnerList.TrimEnd(",")
                        $missingDependenciesObjectArray += $missingDependencyObject
                        #endregion - Build & Add Object to the Array
                        #region - Delete Missing Feature Reference
                        if ($deleteMissingFeatures -eq $True)
                            {
                            $removeFeature = AskYesNoQuestion "   + Do you want to remove this feature reference?"
                            if ($removeFeature -eq $true)
                                {
                                Write-Host ("    - Removing")
                                Remove-SPFeatureFromContentDB -ContentDB $databaseName -FeatureId $featureId –ReportOnly #Remove the REPORTONLY safety to apply changes
                                Write-Host ("...Done") -ForegroundColor Green
                                }
                            else
                                {
                                Write-Host ("    - Skipping...") -ForegroundColor Yellow
                                }
                            }
                        #endregion - Delete Missing Feature Reference
                        }
                    }
                }
            else
                {
                Write-Host (" - Query Returned nothing `"" + $featureId + "`" on Database `"" + $errorDatabase.Name + "`"") -ForegroundColor Yellow
                }            
            }
            #endregion - Iteration
        #endregion - MISSING FEATURE
        #region - MISSING WEBPART
        elseif ($missingDependency.Category -eq "MissingWebPart")
            {
            #region - REGEX
            $startMatches = ([regex]'\[').Matches($missingDependency.Message);            
            $endMatches = ([regex]'\]').Matches($missingDependency.Message);            
            $StartIndex0 = $startMatches[0].Index + 1
            $EndIndex0 = $endMatches[0].Index - $StartIndex0
            $StartIndex1 = $startMatches[$startMatches.Count -2].Index + 1
            $EndIndex1 = $endMatches[$startMatches.Count -2].Index - $StartIndex1
            #endregion - REGEX
            #region - Get Missing Dependency Info
            $webpartClass = $missingDependency.Message.Substring($StartIndex0,$EndIndex0)               
            $webpartOccurences = $missingDependency.Message.Substring($StartIndex1,$EndIndex1)  
            #endregion - Get Missing Dependency Info          
            #region - SQL Query
            $databaseServer = (Get-SPContentDatabase $errorDatabase.Name).Server
            $sqlQueryWebPart = "SELECT DISTINCT AllDocs.Id as DocId, AllDocs.SiteId, AllDocs.WebId, AllDocs.DirName, AllDocs.LeafName, AllDocs.ListId, tp_ZoneID, tp_DisplayName, tp_Class, tp_ID, tp_WebPartIdProperty, tp_PageVersion FROM AllDocs WITH (NoLOCK) INNER JOIN AllWebParts on AllDocs.Id = AllWebParts.tp_PageUrlID WHERE AllWebParts.tp_WebPartTypeID = '$webpartClass' ORDER BY DocId, tp_ID, tp_PageVersion"
            $missingWebParts = Run-SQLQuery -SqlServer $databaseServer -SqlDatabase $errorDatabase.Name -SqlQuery $sqlQueryWebPart                        
            #endregion - SQL Query
            #region - Iteration
            Write-Host ("+ " + $missingDependency.Category + ": [" + $webpartClass.ToUpper() + "] Occurs `"" + $webpartOccurences + "`" times in DB: `"" + $errorDatabase.Name.ToUpper() + "`"") -ForegroundColor Magenta           
            Write-Host (" - [" + ($missingWebParts.tp_class | Select-Object -First 1) + "]") -ForegroundColor DarkGray
            $uniqueLeafMissingWebParts = ($missingWebParts | Where-Object {$_.LeafName -ne $null} | Select-Object -Property SiteId, WebId, LeafName, Dirname -Unique)
            foreach ($uniqueLeaf in $uniqueLeafMissingWebParts)
                {
                $i = 0
                foreach ($leafCount in $missingWebParts | Where-Object {$_.LeafName -ne $null})
                    {
                    
                    if ($uniqueLeaf.LeafName -eq $leafCount.LeafName)
                        {
                        $i++                        
                        }
                    }
                $web = Get-SPWeb -Site $uniqueLeaf.SiteId -Id $uniqueLeaf.WebId
                Write-Host (" - ") -NoNewline -ForegroundColor DarkGray
                $webPartUrl = ($web.Site.WebApplication.Url + $uniqueLeaf.DirName + "/" + $uniqueLeaf.LeafName + "?contents=1")
                Write-Host $webPartUrl -NoNewline -ForegroundColor White
                Write-Host (" - (" + $i + " Occurrences)") -ForegroundColor DarkGray                
                #region - Build Web Owner String
                $webOwnerList = ""
                foreach ($webOwner in $web.AssociatedOwnerGroup.Users)
                    {
                    $webOwnerList += ("`"" + $webOwner.DisplayName + "`",")
                    } 
                #endregion - Build Web Owner String
                #region - Build & Add Object to the Array 
                $missingDependencyObject = New-Object –TypeName PSObject
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Database -Value $errorDatabase.Name.ToUpper()                       
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Category -Value $missingDependency.Category
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Id -Value $missingWebParts.tp_ID
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Name -Value $missingWebParts.tp_class
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Scope -Value "Web" 
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Url -Value $webPartUrl
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Occurences -Value $i
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Owners -Value $webOwnerList.TrimEnd(",")
                $missingDependenciesObjectArray += $missingDependencyObject                  
                #endregion - Build & Add Object to the Array 
                }
            #endregion - Iteration                     
            }
        #endregion - MISSING WEBPART
        #region - MISSING  ASSEMBLY
        elseif ($missingDependency.Category -eq "MissingAssembly")
            {
            #region - REGEX
            $startMatches = ([regex]'\[').Matches($missingDependency.Message);            
            $endMatches = ([regex]'\]').Matches($missingDependency.Message);            
            $StartIndex0 = $startMatches[0].Index + 1
            $EndIndex0 = $endMatches[0].Index - $StartIndex0
            #endregion - REGEX
            #region - Get Missing Dependency Info
            $MissingAssemblyName = $missingDependency.Message.Substring($StartIndex0,$EndIndex0)
            #endregion - Get Missing Dependency Info
            #region - SQL Query
            $databaseServer = (Get-SPContentDatabase $errorDatabase.Name).Server                        
            $sqlQueryAssemblies = "SELECT DISTINCT Id, Name, SiteId, WebId, HostId, HostType FROM EventReceivers WITH (NoLOCK) WHERE Assembly = '$MissingAssemblyName'"
            $missingAssemblies = Run-SQLQuery -SqlServer $databaseServer -SqlDatabase $errorDatabase.Name -SqlQuery $sqlQueryAssemblies                        
            #endregion - SQL Query
            #region - Iteration
            Write-Host ("+ " + $missingDependency.Category + ": [" + $MissingAssemblyName + "]" + " missing (" + ($missingAssemblies.Count -1) + "x) in database `"" + $errorDatabase.Name + "`"") -ForegroundColor Green          
            $i = 1
            foreach ($missingAssembly in $missingAssemblies | Where-Object {$_.Id -ne $null})
                {
                #region - Write Number and Name
                Write-Host (" - [" + $i + "]") -NoNewline -ForegroundColor Gray
                if ($missingAssembly.Name)
                    {
                    Write-Host ("[Name: " + $missingAssembly.Name + "]") -NoNewline -ForegroundColor Gray
                    }
                else
                    {
                    Write-Host ("[Name: None]") -NoNewline -ForegroundColor Gray
                    }
                $i++
                #endregion - Write Number and Name
                #region - Determine the Missing Assembly Scope
                # NOTE: Event Host Type: https://msdn.microsoft.com/en-us/library/ee394866.aspx                
                switch ($missingAssembly.HostType) {
                "-1" {$missingAssemblyScope = "Invalid"; break}
                "0" {$missingAssemblyScope = "Site"; break}
                "1" {$missingAssemblyScope = "Web"; break}
                "2" {$missingAssemblyScope = "List"; break}
                "3" {$missingAssemblyScope = "List Item"; break}
                "4" {$missingAssemblyScope = "Content Type"; break}
                "5" {$missingAssemblyScope = "Workflow"; break}
                "6" {$missingAssemblyScope = "Feature"; break}
                default {$missingAssemblyScope = "Unkown"; break}
                }
                #endregion - Determine the Missing Assembly Scope                
                #region - Get Assembly Web URL
                # TO BE EXPANDED TO DETERMINE THE URL OF EACH OF THE SWITCH CASES ABOVE
                #if ($missingAssembly.WebId -eq "00000000-0000-0000-0000-000000000000")
                if ($missingAssemblyScope -eq "0") #SITE
                    {
                    $site = Get-SPSite -Identity $missingAssembly.SiteId.ToString()
                    $missingAssemblyUrl = $site.Url
                    Write-Host (" - [$missingAssemblyScope Scoped] - URL: ") -NoNewline -ForegroundColor Gray
                    Write-Host $missingAssemblyUrl
                    }
                elseif ($missingAssemblyScope -eq "1") #WEB
                    {
                    $web = Get-SPWeb -Site $missingAssembly.SiteId -Id $missingAssembly.WebId
                    $missingAssemblyUrl = $web.Url
                    Write-Host (" - [$missingAssemblyScope Scoped] " ) -NoNewline -ForegroundColor Gray                          
                    Write-Host $web.Url
                    }
                elseif ($missingAssemblyScope -eq "2") #LIST
                    {
                    $web = Get-SPWeb -Site $missingAssembly.SiteId -Id $missingAssembly.WebId
                    $list = $web.Lists | Where-Object {$_.ID -eq $missingAssembly.HostId}
                    $missingAssemblyUrl = ($web.Url + $list.DefaultViewUrl)
                    Write-Host (" - [$missingAssemblyScope Scoped] " ) -NoNewline -ForegroundColor Gray                          
                    Write-Host $web.Url
                    }
                elseif ($missingAssemblyScope -eq "3") #LIST ITEM
                    {
                    #NOT TESTED
                    $web = Get-SPWeb -Site $missingAssembly.SiteId -Id $missingAssembly.WebId
                    $file = $web.GetFile([Guid]$missingAssembly.HostId)
                    $missingAssemblyUrl = ($web.Url + "/" + $file.Url)
                    Write-Host (" - [$missingAssemblyScope Scoped] " ) -NoNewline -ForegroundColor Gray                          
                    Write-Host $web.Url
                    }
                elseif ($missingAssemblyScope -eq "4") #CONTENT TYPE
                    {
                    #NOT TESTED
                    $web = Get-SPWeb -Site $missingAssembly.SiteId -Id $missingAssembly.WebId
                    $missingAssemblyUrl = ($web.Url + "/" + $file.Url)
                    Write-Host (" - [$missingAssemblyScope Scoped] " ) -NoNewline -ForegroundColor Gray                          
                    Write-Host $web.Url
                    }
                elseif ($missingAssemblyScope -eq "5") #WORKFLOW
                    {
                    #NOT TESTED
                    $web = Get-SPWeb -Site $missingAssembly.SiteId -Id $missingAssembly.WebId
                    $webWorkFlowAssociation = $web.WorkflowAssociations | Where-Object {$_.Id -eq $missingAssembly.HostId}
                    if ($webWorkFlowAssociation)
                        {
                        Write-Host ("Workflow " +  + " on " + $list.DefaultViewUrl)
                        #$workflowFound = $true
                        $missingAssemblyUrl = ($web.Url + " [" + $webWorkFlowAssociation.Name + "]")
                        }
                    else
                        {
                        foreach ($list in $web.Lists)
                            {
                            $listWorkFlowAssociation = $list.WorkflowAssociations | Where-Object {$_.Id -eq $missingAssembly.HostId}
                            if ($listWorkFlowAssociation)
                                {
                                Write-Host ("Workflow `"" + $listWorkFlowAssociation.Name + "`" on list: " + ($site.url + $list.DefaultViewUrl))
                                #$workflowFound = $true 
                                $missingAssemblyUrl = ($web.url + $list.DefaultViewUrl+ " [" + $listWorkFlowAssociation.Name + "]")
                                }
                            else
                                {                       
                                }
                            }
                        if (!$workflowNotFound) {Write-Host ("Could not find Workflow");$missingAssemblyUrl = $web.Url}
                        }
                    
                    Write-Host (" - [$missingAssemblyScope Scoped] " ) -NoNewline -ForegroundColor Gray                          
                    Write-Host $missingAssemblyUrl
                    }
                elseif ($missingAssemblyScope -eq "6") #FEATURE
                    {
                    #NOT TESTED
                    $web = Get-SPWeb -Site $missingAssembly.SiteId -Id $missingAssembly.WebId
                    $feature = $web.Features[$missingAssembly.HostId]                    
                    $missingAssemblyUrl = ($web.Url + " - [" + $feature.DefinitionId.Guid.ToString() + "]")
                    Write-Host (" - [$missingAssemblyScope Scoped] " ) -NoNewline -ForegroundColor Gray                          
                    Write-Host $web.Url
                    }
                else
                    {
                    #NOT TESTED
                    $site = Get-SPSite -Identity $missingAssembly.SiteId.ToString()
                    $missingAssemblyUrl = $site.Url
                    Write-Host (" - [$missingAssemblyScope Scoped] - URL: ") -NoNewline -ForegroundColor Gray
                    Write-Host $missingAssemblyUrl
                    }
                #endregion - Get Assembly Web URL
                #region - Build Web Owner String
                $webOwnerList = ""
                foreach ($webOwner in $web.AssociatedOwnerGroup.Users)
                    {
                    $webOwnerList += ("`"" + $webOwner.DisplayName + "`",")
                    } 
                #endregion - Build Web Owner String
                #region - Build & Add Object to the Array                 
                $missingDependencyObject = New-Object –TypeName PSObject
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Database -Value $errorDatabase.Name.ToUpper()                       
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Category -Value $missingDependency.Category
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Id -Value $missingAssembly.HostId
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Name -Value $MissingAssemblyName
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Scope -Value $missingAssemblyScope
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Url -Value $missingAssemblyUrl
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Occurences -Value $missingAssemblies.Count
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Owners -Value $webOwnerList.TrimEnd(",")
                $missingDependenciesObjectArray += $missingDependencyObject 
                #endregion - Build & Add Object to the Array  
                }  
             #endregion - Iteration          
            }
        #endregion - MISSING  ASSEMBLY
        #region - MISSING  SETUP FILE
        elseif ($missingDependency.Category -eq "MissingSetupFile")
            {
            #region - REGEX
            $startMatches = ([regex]'\[').Matches($missingDependency.Message);            
            $endMatches = ([regex]'\]').Matches($missingDependency.Message);            
            $StartIndex0 = $startMatches[0].Index + 1
            $EndIndex0 = $endMatches[0].Index - $StartIndex0
            $StartIndex1 = $startMatches[1].Index + 1
            $EndIndex1 = $endMatches[1].Index - $StartIndex1
            #endregion - REGEX
            #region - Get Missing Dependency Info
            $MissingSetupFileName = $missingDependency.Message.Substring($StartIndex0,$EndIndex0)
            $MissingSetupFileOccurences = $missingDependency.Message.Substring($StartIndex1,$EndIndex1)
            #endregion - Get Missing Dependency Info
            #region - SQL Query
            $databaseServer = (Get-SPContentDatabase $errorDatabase.Name).Server                        
            $sqlQuerySetupFile = "SELECT DISTINCT Id, SiteId, DirName, LeafName, WebId, ListId from AllDocs WITH (NoLOCK) WHERE SetupPath = '$MissingSetupFileName'"
            $missingSetupFiles = Run-SQLQuery -SqlServer $databaseServer -SqlDatabase $errorDatabase.Name -SqlQuery $sqlQuerySetupFile
            #endregion - SQL Query
            #region - Iteration 
            Write-Host ("+ " + $missingDependency.Category + ": [" + $MissingSetupFile + "]" + " has `"" + $MissingSetupFileOccurences + "`" occurences in database `"" + $errorDatabase.Name + "`"") -ForegroundColor DarkGray
            foreach ($missingSetupFile in ($missingSetupFiles | Where-Object {$_.Id -ne $null}))
                {
                #region - Get Missing Setup File Url
                $web = Get-SPWeb -Site $missingSetupFile.SiteId -Id $missingSetupFile.WebId
                $site = $web.Site
                $missingFile = $web.GetFile([Guid]$missingSetupFile.Id)
                $MissingSetupFileUrl = ($web.Site.WebApplication.Url.TrimEnd("/") + $missingFile.ServerRelativeUrl)
                Write-Host (" - " + $MissingSetupFileUrl) -ForegroundColor Gray                    
                #endregion - Get Missing Setup File Url
                #region - Build Web Owner String
                $webOwnerList = ""
                foreach ($webOwner in $web.AssociatedOwnerGroup.Users)
                    {
                    $webOwnerList += ("`"" + $webOwner.DisplayName + "`",")
                    }  
                #endregion - Build Web Owner String
                #region - Build & Add Object to the Array                  
                $missingDependencyObject = New-Object –TypeName PSObject
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Database -Value $errorDatabase.Name.ToUpper()                       
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Category -Value $missingDependency.Category
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Id -Value $missingSetupFile.Id
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Name -Value $MissingSetupFileName
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Scope -Value "Web"
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Url -Value $MissingSetupFileUrl
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Occurences -Value $MissingSetupFileOccurences
                $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Owners -Value $webOwnerList.TrimEnd(",")
                $missingDependenciesObjectArray += $missingDependencyObject  
                #endregion - Build & Add Object to the Array
                #region - Delete Missing Feature Reference
                if ($deleteMissingSetupFiles -eq $True)
                    {
                    $removeMissingSetupFile = AskYesNoQuestion ("   + Do you want to DELETE `"" + $missingFile.Name + "`" (including all Recycle Bin Stages)? ")
                    if ($removeMissingSetupFile -eq $true)
                        {
                        Write-Host ("Deleting `"" + $missingFile.Name + "`"") -NoNewline
                        $missingFile.Delete()
                        Write-Host ("...Done") -ForegroundColor Green
                        $webRecycleBinItems = $web.RecycleBin | Where-Object {$_.Title -eq $missingFile.Name}
                        if ($null -ne $webRecycleBinItems)
                            {
                            foreach ($webRecycleBinItem in $webRecycleBinItems)
                                {
                                Write-Host ("Deleting `"" + $siteRecycleBinItem.Title + "`" from Web Recycle Bin") -NoNewline
                                $webRecycleBinItem.Delete()
                                Write-Host ("...Done") -ForegroundColor Green
                                }
                            }
                        else
                            {
                            Write-Host ("`"" + $siteRecycleBinItem.Title + "`" not found in Web Recycle Bin") -ForegroundColor Gray
                            }
                        $siteRecycleBinItems = $site.RecycleBin | Where-Object {$_.Title -eq $missingFile.Name}
                        if ($null -ne $siteRecycleBinItems)
                            {
                            foreach ($siteRecycleBinItem in $siteRecycleBinItems)
                                {
                                Write-Host ("Deleting `"" + $siteRecycleBinItem.Title + "`" from Site Collection Recycle Bin") -NoNewline
                                $siteRecycleBinItem.Delete()
                                Write-Host ("...Done") -ForegroundColor Green
                                }    
                            }
                        else
                            {
                            Write-Host ("`"" + $siteRecycleBinItem.Title + "`" not found in Site Collection Recycle Bin") -ForegroundColor Gray
                            }
                        }
                    else
                        {
                        Write-Host ("    - Skipping...") -ForegroundColor Yellow
                        }
                    }
                else
                    {
                    #DO NOTHING
                    }
                #endregion - Delete Missing Feature Reference
                }
            #endregion - Iteration 
            }
        #endregion - MISSING  SETUP FILE
        #region - SITEORPHAN
        elseif ($missingDependency.Category -eq "SiteOrphan")
            {
            #region - REGEX
            $startMatches = ([regex]'\[').Matches($missingDependency.Message);            
            $endMatches = ([regex]'\]').Matches($missingDependency.Message);   
            for ($i = 0; $i -lt $startMatches.Count ; $i++)
                {                
                switch ($i) 
                    {                    
                    "0" {                        
                        $StartIndex = $startMatches[$i].Index + 1
                        $EndIndex = $endMatches[$i].Index - $StartIndex
                        $siteOrphanDatabaseName = ($missingDependency.Message.Substring($StartIndex,$EndIndex))                        
                        break
                        }
                    "1" {
                        $StartIndex = $startMatches[$i].Index + 1
                        $EndIndex = $endMatches[$i].Index - $StartIndex
                        $siteOrphanOrphanedSiteId = ($missingDependency.Message.Substring($StartIndex,$EndIndex))                        
                        break
                        }
                    "2" {
                        $StartIndex = $startMatches[$i].Index + 1
                        $EndIndex = $endMatches[$i].Index - $StartIndex
                        $siteOrphanOrphanedSiteRelativeUrl = ($missingDependency.Message.Substring($StartIndex,$EndIndex))                        
                        break
                        }
                    "3" {
                        $StartIndex = $startMatches[$i].Index + 1
                        $EndIndex = $endMatches[$i].Index - $StartIndex
                        #$siteOrphanOffendingDatabaseId = ($missingDependency.Message.Substring($StartIndex,$EndIndex))                        
                        break
                        }
                    "4" {
                        $StartIndex = $startMatches[$i].Index + 1
                        $EndIndex = $endMatches[$i].Index - $StartIndex
                        $siteOrphanOffendingDatabaseName = ($missingDependency.Message.Substring($StartIndex,$EndIndex))                        
                        break
                        }
                    default 
                        {
                        #$SiteOrphan = "";
                        break;
                        }
                    }
                }            
            #endregion - REGEX
            #region - Get Missing Dependency Info
            $siteOrphanUrl = ((Get-SPContentDatabase $errorDatabase.Name).WebApplication.Url.TrimEnd("/") + $siteOrphanOrphanedSiteRelativeUrl)
            Write-Host ("+ " + $missingDependency.Category + ": Database [" + $siteOrphanDatabaseName + "] has a Orphaned Site [" + $siteOrphanOrphanedSiteRelativeUrl + "] with Id [" + $siteOrphanOrphanedSiteId + "]") -ForegroundColor Yellow -NoNewline
            Write-Host
            #endregion - Get Missing Dependency Info
            #region - Build & Add Object to the Array       
            $missingDependencyObject = New-Object –TypeName PSObject
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Database -Value $siteOrphanDatabaseName                   
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Category -Value $missingDependency.Category
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Id -Value $siteOrphanOrphanedSiteId 
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Name -Value $siteOrphanOffendingDatabaseName
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Scope -Value "Site"
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Url -Value $siteOrphanUrl
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Occurences -Value "Unknown"
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Owners -Value "Unknown"
            $missingDependenciesObjectArray += $missingDependencyObject
            #endregion - Build & Add Object to the Array                                   
            }
        #endregion - SITEORPHAN
        #region - UNHANDLED MISSING DEPENDENCY
        else
            {
            #region - REGEX
            $startMatches = ([regex]'\[').Matches($missingDependency.Message);            
            $endMatches = ([regex]'\]').Matches($missingDependency.Message);   
            $MissingDependencyInfo = ""
            for ($i = 0; $i -lt $startMatches.Count ; $i++)
                {
                $StartIndex = $startMatches[$i].Index + 1
                $EndIndex = $endMatches[$i].Index - $StartIndex
                $MissingDependencyInfo += ("[" + $missingDependency.Message.Substring($StartIndex,$EndIndex) + "]")
                }            
            #endregion - REGEX
            #region - Get Missing Dependency Info
            Write-Host ("+ " + $missingDependency.Category + ": " + $MissingDependencyInfo + " ") -ForegroundColor Red -NoNewline
            Write-Host
            #endregion - Get Missing Dependency Info
            #region - Build & Add Object to the Array                                   
            $missingDependencyObject = New-Object –TypeName PSObject
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Database -Value $errorDatabase.Name.ToUpper()                       
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Category -Value $missingDependency.Category
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Id -Value "Unknown"
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Name -Value $MissingDependencyInfo
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Scope -Value "Unknown"
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Url -Value "Unknown"
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Occurences -Value "Unknown"
            $missingDependencyObject | Add-Member –MemberType NoteProperty –Name Owners -Value "Unknown"
            $missingDependenciesObjectArray += $missingDependencyObject
            #endregion - Build & Add Object to the Array                                   
            }
        #endregion - UNHANDLED MISSING DEPENDENCY
        }
    #endregion - Missing Dependency Iteration
    WriteCharLine -lineFillCharacter $lineFillCharacter -lineLength $lineLength -lineColor $lineColor
    }
#region - Write Report CSV
Write-Host
WriteCharLine -lineFillCharacter $lineFillCharacter -lineLength $lineLength -lineColor $lineColor
if ($null -ne $missingDependenciesObjectArray)
    {    
    $createCSVReport = AskYesNoQuestion ("/!\ Do you want to export this Report to a CSV File? ")
    if ($createCSVReport -eq $true)
        {
        $missingDependenciesCSV = ($ScriptPath + "\MissingDependenciesReport_" + $startDateTime + ".CSV")
        Write-Host (" - Exporting Missing Dependencies Details to CSV: `"" + $missingDependenciesCSV + "`"") -NoNewline -ForegroundColor Yellow
        $missingDependenciesObjectArray | Sort-Object -Property Category,Url -Descending | Export-Csv -LiteralPath $missingDependenciesCSV -NoTypeInformation -Encoding UTF8
        Write-Host (" ...Done") -ForegroundColor Green
        }
    else
        {
        Write-Host ("Skipping CSV Export") -ForegroundColor Yellow
        }
    }
else
    {
    Write-Host (" - Nothing to export to CSV") -ForegroundColor Yellow
    }
WriteCharLine -lineFillCharacter $lineFillCharacter -lineLength $lineLength -lineColor $lineColor
#endregion - Write Report CSV
#endregion - MAIN
###########################################################################################################################################
}
#endregion - TRY
###########################################################################################################################################
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
###########################################################################################################################################
