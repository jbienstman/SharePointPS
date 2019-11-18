Function DatabaseAndSiteCollectionCreator {
<#
.NOTES
    #######################################################################################################################################
    # Author: Jim B.
    # Disclaimer: No Guarantees that this will work on "Any SharePoint Farm"
    #######################################################################################################################################
    #######################################################################################################################################
    # Revisions
    # 1.0 - 2018-06-26 - Initial Commit
    #######################################################################################################################################
.SYNOPSIS
    Creates a Database and Optional "Path Based" Site Collection
.DESCRIPTION
    Creates a Database and Optional "Path Based" Site Collection
.LINK
    None
.EXAMPLE
    ALL PARAMETERS
        DatabaseAndSiteCollectionCreator -WebApplicationName "Content Web Application" -databaseNamePrefix "DBPREFIX_" -DatabaseName "DBNAME1" -CreateSiteCollection $true -ManagedPath "sites" -SiteCollectionName "TEST" -SiteCollectionTemplate "STS#0" -primarySiteCollectionOwnerAlias ($env:USERDOMAIN + "\" + $env:USERNAME)  -MaxSiteCount 1 -WarningSiteCount 0 -SiteLanguage 1033 -SQLAliasName SQLALIAS
#>
Param(
[parameter(mandatory=$false, HelpMessage = "Target Web Application - Name or Url (Will Prompt)")][string]$WebApplicationName ,
[parameter(mandatory=$false, HelpMessage = "Prefix to add to the Database Name ")][string]$databaseNamePrefix = "" ,
[parameter(mandatory=$false, HelpMessage = "Site Language Code: 1033 is default (English)")][int]$SiteLanguage = "1033" ,
[parameter(mandatory=$false, HelpMessage = "New Database Name (check prefix value)(Will Prompt)")][string]$DatabaseName ,
[parameter(mandatory=$false, HelpMessage = "Target SQL Server Alias")][string]$SQLAliasName ,
[parameter(mandatory=$false, HelpMessage = "Maximum Number of Sites on the New Database")][int]$MaxSiteCount = 1 ,
[parameter(mandatory=$false, HelpMessage = "Warning when Site Number reaches this value")][int]$WarningSiteCount = 0 ,
[parameter(mandatory=$false, HelpMessage = "Specify TRUE to create a Site Collection")][bool]$CreateSiteCollection = $false,
[parameter(mandatory=$false, HelpMessage = "Enter valid Managed Path (Will Prompt)")][string]$ManagedPath ,
[parameter(mandatory=$false, HelpMessage = "SiteCollectionName (Will Prompt)")][string]$SiteCollectionName ,
[parameter(mandatory=$false, HelpMessage = "SiteCollectionTemplate")][string]$SiteCollectionTemplate = "STS#0" ,
[parameter(mandatory=$false, HelpMessage = "Primary Site Collection Owner Alias - Default is Current User")][string]$primarySiteCollectionOwnerAlias = ($env:USERDOMAIN + "\" + $env:USERNAME) ,
[parameter(mandatory=$false, HelpMessage = "Secondary Site Collection Owner Alias - Default is Farm Account")][string]$secondarySiteCollectionOwnerAlias ,
[parameter(mandatory=$false, HelpMessage = "Switch to Create default Visitor/Member/Owner Groups")][bool]$createAssociatedSiteGroups = $true
)
$scriptAction = "Create Content Databases and Site Collections"
#region - Output Script Title
Write-Host
Write-Host $scriptAction -ForegroundColor Cyan
for ($i = 0; $i -le $scriptAction.Length; $i++)
    {
    Write-Host "=" -NoNewline -ForegroundColor Cyan
    }
Write-Host
#endregion - Output Script Title
#region - Get Database Information
Write-Host "+ Getting SharePoint Farm Information" -NoNewline
$farm = Get-SPFarm -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
Write-Host ("...Done") -ForegroundColor Green
if (!($WebApplicationName))
    {
    #region - Input
    $menuListOptions = Get-SPWebApplication
    $selection = MenuListChoice $menuListOptions "Choose target Web Application" "Type Number"   
    #endregion - Input
    #region - Output
    $targetWebApplication = Get-SPWebApplication $selection
    $WebApplicationName = $targetWebApplication.Name
    #endregion - Output
    }
else
    {$targetWebApplication = Get-SPWebApplication $WebApplicationName}

if (!($DatabaseName))
    {   
    Write-Host ""
    Write-Host " - Type the name of the Database (e.g: `"" -NoNewline
    Write-Host ($targetWebApplication.ContentDatabases | Select-Object -First 1).Name.ToString() -NoNewline -ForegroundColor Yellow
    Write-Host ("`" NOTE: The prefix [" +  $databaseNamePrefix  + "] is added automatically") -NoNewline
    $DatabaseName = EnsureStringIsTextOnly "!)"
    $DatabaseName = $databaseNamePrefix + $DatabaseName
    Write-Host ("-- The database name will be: ") -NoNewline
    Write-Host ($DatabaseName) -ForegroundColor Yellow
    }
else
    {$DatabaseName = $databaseNamePrefix + $DatabaseName}

#endregion - Get Database Information
#region - Get Site Collection Information
if ($CreateSiteCollection -eq $True)
    {
    if(!($ManagedPath))
        {
        $menuListOptions = Get-SPManagedPath -WebApplication $WebApplicationName | Where-Object {$_.Type -eq "WildcardInclusion"} | Select-Object Name -ExpandProperty Name
        $ManagedPath = MenuListChoice $menuListOptions "Choose target SPManagedPath" "Type Number"
        }
    if (!($siteCollectionName))
        {   
        Write-Host ""
        $siteCollectionName = EnsureStringIsTextOnly " - Type the name of the Site Collection (This will also be the URL)"        
        }

    if ($ManagedPath -eq "/")
        {
        $siteCollectionUrl = $targetWebApplication.Url
        Write-Host ("-- The siteCollectionUrl name will be: ") -NoNewline
        Write-Host ($siteCollectionUrl) -ForegroundColor Yellow
        }
    else
        {
        $siteCollectionUrl = ($targetWebApplication.Url + $ManagedPath).TrimEnd("/") + "/" + $siteCollectionName
        Write-Host ("-- The siteCollectionUrl name will be: ") -NoNewline
        Write-Host ($siteCollectionUrl) -ForegroundColor Yellow
        }
    }
else
    {
    }
#endregion - Get Site Collection Information
#region - EXECUTION
#region - Create Content Database
if (!(Get-SPContentDatabase $DatabaseName -WarningAction SilentlyContinue -ErrorAction SilentlyContinue))
    {
    Write-Host ("+ Creating Database: ") -NoNewline
    Write-Host ($DatabaseName) -ForegroundColor Yellow -NoNewline
    if (!($SQLAliasName))
        {
        New-SPContentDatabase -WebApplication $WebApplicationName -Name $DatabaseName -MaxSiteCount $maxSiteCount -WarningSiteCount $warningSiteCount -Confirm:$false
        }
    else
        {
        New-SPContentDatabase -WebApplication $WebApplicationName -Name $DatabaseName -MaxSiteCount $maxSiteCount -WarningSiteCount $warningSiteCount -DatabaseServer $SQLAliasName -Confirm:$false 
        }    
    Write-Host ("...Done") -ForegroundColor Green
    }
else
    {
    Write-Host ("/!\ Database: ") -NoNewline
    Write-Host $DatabaseName -ForegroundColor DarkGray -NoNewline
    Write-Host (" Already exists!") -ForegroundColor Yellow
    }
#endregion - Create Content Database
#region - Create Site Collection
if ($CreateSiteCollection -eq $true)
    {
    if ($secondarySiteCollectionOwnerAlias -eq "")
        {
        $secondarySiteCollectionOwnerAlias = $farm.DefaultServiceAccount.Name.ToString()
        }
    if (Get-SPContentDatabase $DatabaseName -WarningAction SilentlyContinue -ErrorAction SilentlyContinue)
        {        
        if (!(Get-SPSite $siteCollectionUrl -WarningAction SilentlyContinue -ErrorAction SilentlyContinue))
            {
            Write-Host ("+ Creating Site Collection: ") -NoNewline
            Write-Host ($siteCollectionUrl) -ForegroundColor Yellow -NoNewline
            if ($siteCollectionTemplate -eq "" -or $siteCollectionTemplate -eq "none")
                {
                $newlyCreatedSite = New-SPSite -Url $siteCollectionUrl -Name $siteCollectionName -Language $SiteLanguage -Description $siteCollectionName  -OwnerAlias $primarySiteCollectionOwnerAlias -SecondaryOwnerAlias $secondarySiteCollectionOwnerAlias -ContentDatabase $DatabaseName -Confirm:$false
                }
            else
                {            
                $newlyCreatedSite = New-SPSite -Url $siteCollectionUrl -Template $siteCollectionTemplate -Name $siteCollectionName -Language $SiteLanguage -Description $siteCollectionName  -OwnerAlias $primarySiteCollectionOwnerAlias -SecondaryOwnerAlias $secondarySiteCollectionOwnerAlias -ContentDatabase $DatabaseName -Confirm:$false
                }
            Write-Host ("...Done") -ForegroundColor Green
            }
        else
            {
            Write-Host ("/!\ Site Collection: ") -NoNewline
            Write-Host ($siteCollectionUrl) -ForegroundColor DarkGray -NoNewline            
            Write-Host (" Aleady Exists!") -ForegroundColor Yellow
            }
        }
    else
        {
        Write-Host ("! Database: ") -NoNewline
        Write-Host ($DatabaseName) -ForegroundColor DarkGray -NoNewline
        Write-Host (" does not exist - Skipping Site Creation")
        }   
    #region - Create Associated Groups
    if ($createAssociatedSiteGroups -eq $true)
        {
        if (!$newlyCreatedSite)
            {
            $newlyCreatedSite = Get-SPSite $siteCollectionUrl
            }
        $newlyCreatedRootWeb = $newlyCreatedSite.RootWeb
        if ($newlyCreatedRootWeb.AssociatedVisitorGroup -eq $null -and $newlyCreatedRootWeb.AssociatedMemberGroup -eq $null -and $newlyCreatedRootWeb.AssociatedOwnerGroup -eq $null)
            {
            Write-Host (" + + Adding Default Groups: ") -NoNewline    
            $newlyCreatedRootWeb.CreateDefaultAssociatedGroups($primarySiteCollectionOwnerAlias,$primarySiteCollectionOwnerAlias,$siteCollectionName)  
            Write-Host "..." -NoNewline              
            $newlyCreatedRootWeb.Update()
            Write-Host ("Done") -ForegroundColor Green
            }
        else
            {
            Write-Host (" + + at least one of the Default Visitor/Member/Owner Groups already exists - skipping automatic creation") 
            }
        }                
    else
        {
        #No Associated Visitor/Member/Owner Groups will be created
        }
    #endregion - Create Associated Groups
    }
else
    {
    Write-Host ("No Site Collection will be created...") -ForegroundColor Yellow
    }
#endregion - Create Site Collection
#endregion - EXECUTION
}
