# Copyright (c) 2023 Austin L

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

#region VARIABLES
$global:wordDirtySearch = $null;

$configFilePath = "https://raw.githubusercontent.com/TheRealGoodLivin/SPOA/main/CONFIG.json"
$currentVersion = "1.1"

$setupPath = if (Test-Path -Path $env:OneDrive) { $env:OneDrive + "\Documents\SOPA" } else { $env:UserProfile + "\Documents\SOPA" }
$setupReportPath = $setupPath + "\Reports"
$setupDirtyWordsPath = $setupPath + "\DirtyWords"
$setupDirtyWordsFilePath = $setupDirtyWordsPath + "\DirtyWords.csv"
#endregion

#region FUNCTIONS
Function Format-FileSize() { # https://community.spiceworks.com/topic/1955251-powershell-help
    Param ([int]$size)
    If ($size -gt 1TB) {[string]::Format("{0:0.00} TB", $size / 1TB)}
    ElseIf ($size -gt 1GB) {[string]::Format("{0:0.00} GB", $size / 1GB)}
    ElseIf ($size -gt 1MB) {[string]::Format("{0:0.00} MB", $size / 1MB)}
    ElseIf ($size -gt 1KB) {[string]::Format("{0:0.00} KB", $size / 1KB)}
    ElseIf ($size -gt 0) {[string]::Format("{0:0.00} B", $size)}
    Else {""}
}

function reportCreate {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][object[]]$reportData)

    if (test-path $reportPath) {
        $reportData | export-csv -Path $reportPath -Force -NoTypeInformation -Append
    } else {
        $reportData | export-csv -Path $reportPath -Force -NoTypeInformation
    }
}
#endregion

#region SETUP FUNCTIONS
function showSetup {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$configFilePath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$currentVersion,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$setupPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$dirtyWordsPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$dirtyWordsFilePath)

    Clear-Host

    try {
        $config = (New-Object System.Net.WebClient).DownloadString($configFilePath) | ConvertFrom-Json
    } catch {
        $config = New-Object PSObject -Property @{
            DirtyWords = @("\d{3}-\d{3}-\d{4}","\d{3}-\d{2}-\d{4}","MyFitness","CUI","UPMR","SURF","PA","2583","SF86","SF 86","FOUO","GTC","medical","AF469","AF 469","469","Visitor Request","VisitorRequest","Visitor","eQIP","EPR","910","AF910","AF 910","911","AF911","AF 911","OPR","eval","feedback","loc","loa","lor","alpha roster","alpha","roster","recall","SSN","SSAN","AF1466","1466","AF 1466","AF1566","AF 1566","1566","SGLV","SF182","182","SF 182","allocation notice","credit","allocation","2583","AF 1466","AF1466","1466","AF1566","AF 1566","1566","AF469","AF 469","469","AF 422","AF422","422","AF910","AF 910","910","AF911","AF 911","911","AF77","AF 77","77","AF475","AF 475","475","AF707","AF 707","707","AF709","AF 709","709","AF 724","AF724","724","AF912","AF 912","912","AF 931","AF931","931","AF932","AF 932","932","AF948","AF 948","948","AF 3538","AF3538","3538","AF3538E","AF 3538E","AF2096","AF 2096","2096","AF 2098","AF2098","AF 2098","AF 3538","AF3538","3538","1466","1566","469","422","travel","SF128","SF 128","128","SF 86","SF86","86","SGLV","SGLI","DD214","DD 214","214","DD 149","DD149","149")
        }
    }
    
    $pnpIsInstalled = Get-InstalledModule -Name PnP.PowerShell -ErrorAction silentlycontinue
    if($pnpIsInstalled.count -eq 0) {
        $Confirm = read-host "`nWOULD YOU LIKE TO INSTALL SHAREPOINT PNP MODULE? [Y] Yes [N] No"
        if($Confirm -match "[yY]") {
            #install-module -Name PnP.PowerShell -scope currentuser
            install-module -Name PnP.PowerShell -RequiredVersion 1.12.0 -Force -scope currentuser #according to PnP: https://github.com/pnp/powershell
        } else {
            write-host "`nSHAREPOINT PNP MODULE IS NEED TO PERFORM THE FEATURES IN THIS SCRIPT." -ForegroundColor red
            break
        }
    } 
    #else {
    #    $pnpCurrentModule = ((get-module -Name PnP.PowerShell -listavailable).Version | sort-object -Descending | select-object -First 1).ToString()
    #    $pnpNewestModule = (find-module -Name PnP.PowerShell).Version.ToString()
    #
    #    if ([System.Version]$pnpCurrentModule -lt [System.Version]$pnpNewestModule) {
    #        $Confirm = read-host "`nTHERE IS AN UPDATE TO SHAREPOINT PNP MODULE. WOULD YOU LIKE TO INSTALL IT? [Y] Yes [N] No"
    #        if($Confirm -match "[yY]") {
    #            update-module -Name PnP.PowerShell
    #        }
    #    }
    #}

    #FOLDER AND FILES SETUP
    if (-Not (test-path $setupPath)) { New-Item -Path $setupPath -ItemType Directory | out-null }
    if (-Not (test-path $reportPath)) { New-Item -Path $reportPath -ItemType Directory | out-null }
    if (-Not (test-path $dirtyWordsPath)) {New-Item -Path $dirtyWordsPath -ItemType Directory | out-null }
    if (-Not (test-path $dirtyWordsFilePath)) { $config.DirtyWords | Select-Object @{Name='Word';Expression={$_}} | export-csv $dirtyWordsFilePath -NoType }
    if (test-path $dirtyWordsPath) { $global:wordDirtySearch = Import-Csv $dirtyWordsFilePath }

    Clear-Host

    #CHECK SPOA UPDATE FILE
    #if ($currentVersion -ne $config.version) {
    #    write-host "###########################################################" -ForegroundColor Green
    #    write-host "#                 NEW SPOA UPDATE AVAIABLE                #" -ForegroundColor Green
    #    write-host "#        https://github.com/TheRealGoodLivin/SPOA/        #" -ForegroundColor Green
    #    write-host "###########################################################`n" -ForegroundColor Green
    #}

    showMain
}
#endregion

#region MAIN AND SETTING MENU FUNCTIONS
function showMain {
    write-host "###########################################################"
    write-host "#                                                         #"
    write-host "#          " -NoNewline
    write-host "  ██████  ██▓███   ▒█████   ▄▄▄      " -ForegroundColor red -NoNewline
    write-host "          #"
    write-host "#          " -NoNewline
    write-host "▒██    ▒ ▓██░  ██▒▒██▒  ██▒▒████▄    " -ForegroundColor red -NoNewline
    write-host "          #"
    write-host "#          " -NoNewline
    write-host "░ ▓██▄   ▓██░ ██▓▒▒██░  ██▒▒██  ▀█▄  " -ForegroundColor red -NoNewline
    write-host "          #"
    write-host "#          " -NoNewline
    write-host "  ▒   ██▒▒██▄█▓▒ ▒▒██   ██░░██▄▄▄▄██ " -ForegroundColor red -NoNewline
    write-host "          #"
    write-host "#          " -NoNewline
    write-host "▒██████▒▒▒██▒ ░  ░░ ████▓▒░ ▓█   ▓██▒" -ForegroundColor red -NoNewline
    write-host "          #"
    write-host "#          " -NoNewline
    write-host "▒ ▒▓▒ ▒ ░▒▓▒░ ░  ░░ ▒░▒░▒░  ▒▒   ▓▒█░" -ForegroundColor red -NoNewline
    write-host "          #"
    write-host "#          " -NoNewline
    write-host "░ ░▒  ░ ░░▒ ░       ░ ▒ ▒░   ▒   ▒▒ ░" -ForegroundColor red -NoNewline
    write-host "          #"
    write-host "#          " -NoNewline
    write-host "░  ░  ░  ░░       ░ ░ ░ ▒    ░   ▒   " -ForegroundColor red -NoNewline
    write-host "          #"
    write-host "#          " -NoNewline
    write-host "      ░               ░ ░        ░  ░" -ForegroundColor red -NoNewline
    write-host "          #"
    write-host "#                                                         #"
    write-host "#     WELCOME TO THE SHAREPOINT ONLINE ASSISTANT TOOL     #"
    write-host "#                                                         #"
    write-host "###########################################################"
}

function showMenu {
    write-host "`nMAIN MENU -- SELECT A CATEGORY`n
`t1: PRESS '1' FOR SITE TOOLS.
`t2: PRESS '2' FOR USER TOOLS.
`t3: PRESS '3' FOR LIST TOOLS.
`t4: PRESS '4' FOR DOCUMENT TOOLS.
`tS: PRESS 'S' FOR SETTINGS.
`tQ: PRESS 'Q' TO QUIT.`n
FOR HELP TYPE NUMBER AND ADD ? (E.G. 1?)`n"
}

function showSettings {   
    write-host "`nSETTINGS -- SELECT AN OPTION`n
`t1: PRESS '1' TO OPEN SPOA FOLDER.
`t2: PRESS '2' TO OPEN THE DIRTY WORD LIST.
`tE: PRESS 'E' TO EXIT BACK TO THE MAIN MENU.`n"
}
#endregion

#region SITE TOOLS FUNCTIONS
function showSiteTools {   
    write-host "`nSITE TOOLS -- SELECT AN OPTION`n
`t1: PRESS '1': SITE MAP REPORT.
`t2: PRESS '2': SITE PII SCAN REPORT.
`t3: PRESS '3': SITE COLLECTION ADMIN REPORT.
`t4: PRESS '4': SITE ADD COLLECTION ADMIN.
`t4: PRESS '5': SITE DELETE COLLECTION ADMIN.
`t5: PRESS '6': SITE COLLECTION GROUP REPORT.
`tE: PRESS 'E' TO EXIT BACK TO THE MAIN MENU.`n
FOR HELP TYPE NUMBER AND ADD ? (E.G. 1?)`n"
}

#region SITE TOOLS FUNCTIONS OPTION "1"
function spoSiteMap {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = read-host "`nENTER SITE COLLECTION URL"
    $results = @()

    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue

    $siteInfo = Get-PnPWeb -Includes Created | select Title, ServerRelativeUrl, Url, Created, Description
    $siteLists = Get-PnPList | where-object {$_.Hidden -eq $false}
    $subSites = Get-PnPSubWeb -Recurse | select Title, ServerRelativeUrl, Url, Created, Description

    $siteListCount = @()
    $siteItemCount = 0
    foreach ($list in $subSiteLists) {
        $siteListCount += $list
        $siteItemCount = $siteItemCount + $list.ItemCount
    }

    # GET PARENT SITE INFO AND LIST COUNT
    $results = New-Object PSObject -Property @{
        Title = $siteInfo.Title
        ItemCount = $siteItemCount
        ListCount = $siteListCount.count
        ServerRelativeUrl = $siteInfo.ServerRelativeUrl
        Description = $siteInfo.Description
        Created = $siteInfo.Created
    }
    reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results

    foreach ($site in $subSites) {
        Connect-PnPOnline -Url $site.Url -UseWebLogin -WarningAction SilentlyContinue
        $subSiteLists = Get-PnPList | where-object {$_.Hidden -eq $false}

        $subSiteListCount = @()
        $subSiteItemCount = 0
        foreach ($list in $subSiteLists) {
            $subSiteListCount += $list
            $siteListCount += $list
            $subSiteItemCount = $subSiteItemCount + $list.ItemCount
            $siteItemCount = $siteItemCount + $list.ItemCount
        }

        $results = New-Object PSObject -Property @{
            Title = $site.Title
            ListCount = $subSiteListCount.count
            ItemCount = $subSiteItemCount
            ServerRelativeUrl = $site.ServerRelativeUrl
            Description = $site.Description
            Created = $site.Created
        }
        reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results
    }

    # GET TOTAL COUNTS
    $results = New-Object PSObject -Property @{
        Title = "Total"
        ListCount = $siteListCount.count
        ItemCount = $siteItemCount
        ServerRelativeUrl = $subSites.count + 1
        Description = ""
        Created = ""
    }
    reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results

    Disconnect-PnPOnline
    write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
}
#endregion

#region SITE TOOLS FUNCTIONS OPTION "2"
function spoSiteScanPII {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = read-host "`nENTER SITE COLLECTION URL"
    $results = @()

    $Confirm = read-host "`nWOULD YOU LIKE TO SCAN ALL SUB-SITES? [Y] Yes [N] No"
    if($Confirm -match "[yY]") {
        $siteParentOnly = $false
    } else {
        $siteParentOnly = $true
    }

    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue
    $getDocLibs = Get-PnPList | where-object { $_.BaseTemplate -eq 101 }

    write-host "Searching: $($sitePath)" -ForegroundColor Green

    foreach ($DocLib in $getDocLibs) {
        Get-PnPListItem -List $DocLib -Fields "FileRef", "File_x0020_Type", "FileLeafRef", "File_x0020_Size", "Created", "Modified" -PageSize 1000 | Where { $_["FileLeafRef"] -like "*.*" } | foreach-object {
            foreach ($word in $global:wordDirtySearch) {
                $wordSearch = "(?i)\b$($word.Word)\b"

                if (($_["FileLeafRef"] -match $wordSearch)) {
                    write-host "File found. " -ForegroundColor Red -nonewline; write-host "Under: '$($word.Word)' Path: $($_["FileRef"])" -ForegroundColor Yellow;

                    $permissions = @()
                    $perm = Get-PnPProperty -ClientObject $_ -Property RoleAssignments       
                    foreach ($role in $_.RoleAssignments) {
                        $loginName = Get-PnPProperty -ClientObject $role.Member -Property LoginName
                        $rolebindings = Get-PnPProperty -ClientObject $role -Property RoleDefinitionBindings
                        $permissions += "$($loginName) - $($rolebindings.Name)"
                    }
                    $permissions = $permissions | Out-String

                    if ($_ -eq $null) {
                        write-host "Error: 'Unable to pull file information'."
                    } else {
                        $size = Format-FileSize($_["File_x0020_Size"])
                               
                        $results = New-Object PSObject -Property @{
                            FileName = $_["FileLeafRef"]
                            FileExtension = $_["File_x0020_Type"]
                            FileSize = $size
                            Path = $_["FileRef"]
                            Permissions = $permissions
                            Criteria = $word.Word
                            Created = $_["Created"]
                            Modified = $_["Modified"]
                        }
                        reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results
                    }
                }
            }
        }
    }

    if ($siteParentOnly -eq $false) {
        $subSites = Get-PnPSubWeb -Recurse

        foreach ($site in $subSites) {
            Connect-PnPOnline -Url $site.Url -UseWebLogin -WarningAction SilentlyContinue
            $getSubDocLibs = Get-PnPList | where-object {$_.BaseTemplate -eq 101}

            write-host "Searching: $($site.Url)" -ForegroundColor Green

            foreach ($subDocLib in $getSubDocLibs) {
                Get-PnPListItem -List $subDocLib -Fields "FileRef", "File_x0020_Type", "FileLeafRef", "File_x0020_Size", "Created", "Modified" -PageSize 1000 | Where { $_["FileLeafRef"] -like "*.*" } | foreach-object {
                    foreach ($word in $global:wordDirtySearch) {
                        $wordSearch = "(?i)\b$($word.Word)\b"

                        if (($_["FileLeafRef"] -match $wordSearch)) {
                            write-host "File found. " -ForegroundColor Red -nonewline; write-host "Under: '$($word.Word)' Path: $($_["FileRef"])" -ForegroundColor Yellow;

                            $permissions = @()
                            $perm = Get-PnPProperty -ClientObject $_ -Property RoleAssignments       
                            foreach ($role in $_.RoleAssignments) {
                                $loginName = Get-PnPProperty -ClientObject $role.Member -Property LoginName
                                $rolebindings = Get-PnPProperty -ClientObject $role -Property RoleDefinitionBindings
                                $permissions += "$($loginName) - $($rolebindings.Name)" 
                            }
                            $permissions = $permissions | Out-String

                            if ($_ -eq $null) {
                                write-host "Error: 'Unable to pull file information'."
                            } else {
                                $size = Format-FileSize($_["File_x0020_Size"])
           
                                $results = New-Object PSObject -Property @{
                                    FileName = $_["FileLeafRef"]
                                    FileExtension = $_["File_x0020_Type"]
                                    FileSize = $size
                                    Path = $_["FileRef"]
                                    Permissions = $permissions
                                    Criteria = $word.Word
                                    Created = $_["Created"]
                                    Modified = $_["Modified"]
                                }
                                reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results
                            }
                        }
                    }
                }
            }
        }
    }

    Disconnect-PnPOnline
    write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
}
#endregion

#region SITE TOOLS FUNCTIONS OPTION "3"
function spoSiteGetCollectionAdmins {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)
    
    $results = @()
    $sitePath = read-host "`nENTER SITE COLLECTION URL"
    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue

    Get-PnPSiteCollectionAdmin | foreach-object {
        $results = New-Object PSObject -Property @{
            Id = $_.Id
            Title = $_.Title
            Email = $_.Email
            LoginName = $_.LoginName
            IsSiteAdmin = $_IsSiteAdmin
            IsShareByEmailGuestUser = $_.IsShareByEmailGuestUser
            IsHiddenInUI = $_.IsHiddenInUI
            PrincipalType = $_.PrincipalType
        }
        reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results
    }

    Disconnect-PnPOnline
    write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
}
#endregion

#region SITE TOOLS FUNCTIONS OPTION "4"
function spoSiteAddCollectionAdmin {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)
    
    $results = @()
    $sitePath = read-host "`nENTER SITE COLLECTION URL"
    $newAdmin = read-host "`nENTER NEW SITE COLLECTION ADMIN EMAIL"
    
    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue
    Add-PnPSiteCollectionAdmin -Owners $newAdmin

    $results = New-Object PSObject -Property @{
        AdminNew = $newAdmin
    }
    reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results

    Disconnect-PnPOnline
    write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
}
#endregion

#region SITE TOOLS FUNCTIONS OPTION "5"
function spoSiteDeleteCollectionAdmin {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)
    
    $results = @()
    $sitePath = read-host "`nENTER SITE COLLECTION URL"
    
    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue

    $getAdmins = @()
    Get-PnPSiteCollectionAdmin | foreach-object { $getAdmins += $_ }

    do {
        write-host "`nPLEASE SELECT AN ADMIN`n"
        foreach ($admin in $getAdmins) {
            write-host "`t$($getAdmins.IndexOf($admin)+1): PRESS $($getAdmins.IndexOf($admin)+1) for $($admin.Title)"
        }
        $adminChoice = read-host "PLEASE MAKE A SELECTION"
    } while (-not($getAdmins[$adminChoice-1]))

    Remove-PnPSiteCollectionAdmin -Owners $getAdmins[$adminChoice-1].Title

    $results = New-Object PSObject -Property @{
        AdminDeleted = $getAdmins[$adminChoice-1].Title
    }
    reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results

    Disconnect-PnPOnline
    write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
}
#endregion

#region SITE TOOLS FUNCTIONS OPTION "6"
function spoSiteGetCollectionGroups {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = read-host "`nENTER SITE COLLECTION URL"
    $results = @()

    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue
    Get-PnPGroup | Where {$_.IsHiddenInUI -eq $false -and $_.LoginName -notlike "Limited Access*" -and $_.LoginName -notlike "SharingLinks*"} | Select-Object "Id", "Title", "LoginName", "OwnerTitle" | foreach-object {
        $members = @()
        Get-PnPGroupMember -Identity $_.Title | foreach-object {
            $members += "$($_.Title)" 
        }
        $members = $members | Out-String

        $results = New-Object PSObject -Property @{
            ID = $_.Id
            GroupName = $_.Title
            LoginName = $_.LoginName
            OwnerTitle = $_.OwnerTitle
            Members = $members
        }
        reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results
    }

    Disconnect-PnPOnline
    write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
}
#endregion
#endregion

#region USER TOOLS FUNCTIONS
function showUserTools {   
    write-host "`nUSER TOOLS -- SELECT AN OPTION`n
`t1: PRESS '1': USER DELETION.
`t2: PRESS '2': USER GROUP DELETION.
`tE: PRESS 'E' TO EXIT BACK TO THE MAIN MENU.`n
FOR HELP TYPE NUMBER AND ADD ? (E.G. 1?)`n"
}

#region USER TOOLS FUNCTIONS OPTION "1"
function spoUserDelete {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = read-host "`nENTER SITE COLLECTION URL"
    $userEmail = read-host "`nENTER USERS EMAIL"
    $results = @()

    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue
    $userInformation = Get-PnPUser | ? Email -eq $userEmail | foreach-object { 
        Remove-PnPUser -Identity $_.Title -Force
        write-host "User Deleted: $($_.Title)" -ForegroundColor Yellow

        $results = New-Object PSObject -Property @{
            UserDeleted = $_.Title
            UserEmail = $_.Email
        }
        reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results
    }

    Disconnect-PnPOnline
    write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
}
#endregion

#region USER TOOLS FUNCTIONS OPTION "2"
function spoUserDeleteGroups {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = read-host "`nENTER SITE COLLECTION URL"
    $userEmail = read-host "`nENTER USERS EMAIL"
    $results = @()

    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue
    $userInformation = Get-PnPUser | ? Email -eq $userEmail | foreach-object { $_.Title }
    $userGroups = Get-PnPUser | ? Email -eq $userEmail | Select -ExpandProperty Groups | Where { ($_.Title -notmatch "Limited Access*") -and ($_.Title -notmatch "SharingLinks*") } | foreach-object { 
        write-host "Name: $userInformation | Group Removed: " -ForegroundColor Yellow -NoNewline; write-host $($_.Title) -ForegroundColor Cyan

        Remove-PnPGroupMember -LoginName $userEmail -Identity $_.Title 

        $results = New-Object PSObject -Property @{
            UserDisplay = $userInformation
            UserGroup = $_.Title
        }
        reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results
    }

    Disconnect-PnPOnline
    write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
}
#endregion
#endregion

#region LIST TOOLS FUNCTIONS
function showListTools {   
    write-host "`nCUSTOM LIST TOOLS -- SELECT AN OPTION`n
`t1: PRESS '1': LIST SHOW IN BROWSER.
`t2: PRESS '2': LIST HIDE FROM BROWSER.
`t3: PRESS '3': LIST DELETE ALL UNIQUE PERMISSIONS.
`t4: PRESS '4': LIST DELETE ALL ITEMS.
`tE: PRESS 'E' TO EXIT BACK TO THE MAIN MENU.`n
FOR HELP TYPE NUMBER AND ADD ? (E.G. 1?)`n"
}

#region LIST TOOLS FUNCTIONS OPTION "1"
function spoListShow {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = read-host "`nENTER SITE URL THAT LIST RESIDES ON"
    $results = @()
    
    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue

    $listsGet = @()
    Get-PnPList | where-object { $_.Hidden -eq $true -and ($_.BaseTemplate -eq 100 -or $_.BaseTemplate -eq 101 -or $_.BaseTemplate -eq 102 -or $_.BaseTemplate -eq 103 -or $_.BaseTemplate -eq 104 -or $_.BaseTemplate -eq 105 -or $_.BaseTemplate -eq 106 -or $_.BaseTemplate -eq 107 -or $_.BaseTemplate -eq 108 -or $_.BaseTemplate -eq 109) } | foreach-object { $listsGet += $_ }

    if ($listsGet.count) {
        do {
            write-host "`nPLEASE SELECT A LIST`n"
            foreach ($list in $listsGet) {
                write-host "`t$($listsGet.IndexOf($list)+1): PRESS $($listsGet.IndexOf($list)+1) for $($list.Title)"
            }
            $listChoice = read-host "`nPLEASE MAKE A SELECTION"
        } while (-not($listsGet[$listChoice-1]))

        Set-PnPList -Identity $listsGet[$listChoice-1].Title -Hidden $false

        $results = New-Object PSObject -Property @{
            ShowList = $listsGet[$listChoice-1].Title
        }
        reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results

        Disconnect-PnPOnline
        write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
        write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
    } else {
        write-host "`nNO LISTS FOUND." -ForegroundColor Red
    }
}
#endregion

#region LIST TOOLS FUNCTIONS OPTION "2"
function spoListHide {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = read-host "`nENTER SITE URL THAT LIST RESIDES ON"
    $results = @()
    
    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue

    $listsGet = @()
    Get-PnPList | where-object { $_.Hidden -eq $false } | foreach-object { $listsGet += $_ }

    if ($listsGet.count) {
        do {
            write-host "`nPLEASE SELECT A LIST`n"
            foreach ($list in $listsGet) {
                write-host "`t$($listsGet.IndexOf($list)+1): PRESS $($listsGet.IndexOf($list)+1) for $($list.Title)"
            }
            $listChoice = read-host "`nPLEASE MAKE A SELECTION"
        } while (-not($listsGet[$listChoice-1]))

        Set-PnPList -Identity $listsGet[$listChoice-1].Title -Hidden $true

        $results = New-Object PSObject -Property @{
            HideList = $listsGet[$listChoice-1].Title
        }
        reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results

        Disconnect-PnPOnline
        write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
        write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
    } else {
        write-host "`nNO LISTS FOUND." -ForegroundColor Red
    }
}
#endregion

#region LIST TOOLS FUNCTIONS OPTION "3"
function spoListDeleteAllUniquePermissions {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = read-host "`nENTER SITE URL THAT LIST RESIDES ON"
    $results = @()
    
    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue

    $listsGet = @()
    Get-PnPList | where-object { $_.Hidden -eq $false } | foreach-object { $listsGet += $_ }

    if ($listsGet.count) {
        do {
            write-host "`nPLEASE SELECT A LIST`n"
            foreach ($list in $listsGet) {
                write-host "`t$($listsGet.IndexOf($list)+1): PRESS $($listsGet.IndexOf($list)+1) for $($list.Title)"
            }
            $listChoice = read-host "`nPLEASE MAKE A SELECTION"
        } while (-not($listsGet[$listChoice-1]))

        $listItems = Get-PnPListItem -List $listsGet[$listChoice-1].Title -PageSize 500
        ForEach($item in $listItems) {
            $checkItemPermissions = Get-PnPProperty -ClientObject $item -Property "HasUniqueRoleAssignments"
            If($checkItemPermissions) {
                Set-PnPListItemPermission -List $listsGet[$listChoice-1].Title -Identity $item.Id -InheritPermissions

                $results = New-Object PSObject -Property @{
                    ListName = $listsGet[$listChoice-1].Title
                    ItemID = $item.Id
                }
                reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results
            }
        }

        Disconnect-PnPOnline
        write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
        write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
    } else {
        write-host "`nNO LISTS FOUND." -ForegroundColor Red
    }
}
#endregion

#region LIST TOOLS FUNCTIONS OPTION "4"
function spoListDeleteAllItems {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = read-host "`nENTER SITE URL THAT LIST RESIDES ON"
    $results = @()
    
    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue
    $listsGet = @()

    Get-PnPList | where-object { $_.Hidden -eq $false } | foreach-object { $listsGet += ($_) }

    if ($listsGet.count) {
        do {
            write-host "`nPLEASE SELECT A LIST`n"
            foreach ($list in $listsGet) {
                write-host "`t$($listsGet.IndexOf($list)+1): PRESS $($listsGet.IndexOf($list)+1) for $($list.Title)"
            }
            $listChoice = read-host "`nPLEASE MAKE A SELECTION"
        } while (-not($listsGet[$listChoice-1]))

        $listItems =  Get-PnPListItem -List $listsGet[$listChoice-1].Title -PageSize 500
        $Batch = New-PnPBatch
        ForEach($item in $listItems) {    
             Remove-PnPListItem -List $listsGet[$listChoice-1].Title -Identity $item.Id -Recycle -Batch $Batch

            $results = New-Object PSObject -Property @{
                ListName = $listsGet[$listChoice-1].Title
                ItemDeletedID = $item.Id
            }
            reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results
        }
        Invoke-PnPBatch -Batch $Batch

        Disconnect-PnPOnline
        write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
        write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
    } else {
        write-host "`nNO LISTS FOUND." -ForegroundColor Red
    }
}
#endregion
#endregion

#region DOCUMENT TOOLS FUNCTIONS
function showDocumentTools {   
    write-host "`nDOCUMENT TOOLS -- SELECT AN OPTION`n
`t1: PRESS '1': DOCUMENT FOLDER UPLOAD.
`t2: PRESS '2': DOCUMENT SHARED LINKS REPORT.
`t3: PRESS '3': DOCUMENT REMOVE ALL SHARED LINKS..
`tE: PRESS 'E' TO EXIT BACK TO THE MAIN MENU.`n
FOR HELP TYPE NUMBER AND ADD ? (E.G. 1?)`n"
}

#region DOCUMENT TOOLS FUNCTIONS OPTION "1"
function spoDocumentFolderUpload {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $results = @()
    $sitePath = read-host "ENTER SITE URL THAT DOCUMENT LIBRARY RESIDES ON"
    $sitePath = $sitePath.Trim(" ", "/")
    $localPath = read-host "ENTER LOCAL DIRECTORY LOCATION TO COPY"
    $selectedLibraryFolder = ""

    $getDocumentLibraries = @()
    
    if ((Get-Item $localPath) -is [System.IO.DirectoryInfo]) {
        Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue

        Get-PnPList | where-object { $_.Hidden -eq $false -and $_.BaseTemplate -eq 101 -and $_.Title -ne "SiteCollectionDocuments" -and $_.Title -ne "Style Library" -and $_.Title -ne "FormServerTemplates" -and $_.Title -ne "Form Templates" } | foreach-object { $getDocumentLibraries += $_ }

        do {
            write-host "`nPLEASE SELECT A DOCUMENT LIBRARY`n"
            foreach ($documentLibrary in $getDocumentLibraries) {
                write-host "`t$($getDocumentLibraries.IndexOf($documentLibrary)+1): PRESS $($getDocumentLibraries.IndexOf($documentLibrary)+1) for $($documentLibrary.Title)"
            }
            $documentLibraryChoice = read-host "`nPLEASE MAKE A SELECTION"
        } while (-not($getDocumentLibraries[$documentLibraryChoice-1]))

        $selectedLibraryURLFolder = $getDocumentLibraries[$documentLibraryChoice-1].RootFolder.ServerRelativeUrl.replace($getDocumentLibraries[$documentLibraryChoice-1].ParentWebUrl,"")

        do {
            $selectedSubFolders = @()
            Get-PnPFolderItem -FolderSiteRelativeUrl $selectedLibraryURLFolder -ItemType Folder | Where { $_.Name -ne "Forms" } | foreach-object { $selectedSubFolders += $_ }

            if($selectedSubFolders.count) {
                write-host "`nPLEASE SELECT A FOLDER TO COPY TO`n"
                foreach ($child in $selectedSubFolders) {
                    write-host "$($selectedSubFolders.IndexOf($child)+1): PRESS $($selectedSubFolders.IndexOf($child)+1) for $($child.Name)"
                }
                write-host "S: PRESS S to Select Current Folder"
                $folderChoice = read-host "`nPLEASE MAKE A SELECTION"
            } else { $folderChoice = "S" }

            if($folderChoice -ne "S") {
                if(-not($selectedSubFolders[$folderChoice-1])) {
                } else {
                    $selectedLibraryURLFolder += "/$($selectedSubFolders[$folderChoice-1].Name)"
                }
            } else {
                $selectedLibraryFolder = $selectedLibraryURLFolder.Trim(" ", "/")
            }
        } while ($selectedLibraryFolder -eq "")

        $Confirm = read-host "WOULD YOU LIKE TO UPLOAD DOCUMENTS TO THIS FOLDER: $($selectedLibraryFolder)? [Y] Yes [N] No"
        if($Confirm -match "[yY]") {
            write-host "`nProcessing Folder: $($localPath)" -f Yellow
            Resolve-PnPFolder -SiteRelativePath $selectedLibraryFolder | out-null    

            $files = Get-ChildItem -Path $localPath -File
            foreach ($file in $files) {
                Add-PnPFile -Path "$($file.Directory)\$($file.Name)" -Folder $selectedLibraryFolder -Values @{"Title" = $($file.Name)} | out-null
                write-host "`tUploaded File: $($file.FullName)" -f Green

                $results = New-Object PSObject -Property @{
                    Type = "File"
                    OriginalLocation = $file.FullName
                    NewLocation = "$($sitePath)/$selectedLibraryFolder/$($file.Name)"
                }
                reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results
            }

            Get-ChildItem -Path $localPath -Recurse -Directory | foreach-object {
                $folderToUpload = ($selectedLibraryFolder+$_.FullName.Replace($localPath,"")).Replace("\","/")

                write-host "Processing Folder: $($_.FullName)" -ForegroundColor Yellow
                Resolve-PnPFolder -SiteRelativePath $folderToUpload | out-null

                $results = New-Object PSObject -Property @{
                    Type = "Folder"
                    OriginalLocation = $_.FullName
                    NewLocation = "$($sitePath)/$($folderToUpload)"
                }
                reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results

                $files = Get-ChildItem -Path $_.FullName -File
                foreach ($file in $files) {
                    Add-PnPFile -Path "$($file.Directory)\$($file.Name)" -Folder $folderToUpload -Values @{"Title" = $($file.Name)} | out-null
                    write-host "`tUploaded File: $($file.FullName)" -ForegroundColor Green

                    $results = New-Object PSObject -Property @{
                        Type = "File"
                        OriginalLocation = $file.FullName
                        NewLocation = "$($sitePath)/$($folderToUpload)/$($file.Name)"
                    }
                    reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results
                }
            }
        }

        Disconnect-PnPOnline
        write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
        write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
    } else {
        write-host "`nPATH SUPPLIED WAS NOT A FOLDER! PLEASE CHECK YOUR LOCAL DIRECTORY PATH AND TRY AGAIN!" -ForegroundColor Red
    }
}
#endregion

#region DOCUMENT TOOLS FUNCTIONS OPTION "2"
function spoDocumentSharedLinks {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = read-host "`nENTER SITE URL THAT LIST RESIDES ON"
    $results = @()
    
    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue

    $listsGet = @()
    Get-PnPList | where-object { $_.Hidden -eq $false -and $_.BaseTemplate -eq 101 } | foreach-object { $listsGet += $_ }

    if ($listsGet.count) {
        do {
            write-host "`nPLEASE SELECT A LIST`n"
            foreach ($list in $listsGet) {
                write-host "`t$($listsGet.IndexOf($list)+1): PRESS $($listsGet.IndexOf($list)+1) for $($list.Title)"
            }
            $listChoice = read-host "`nPLEASE MAKE A SELECTION"
        } while (-not($listsGet[$listChoice-1]))

        $listItems = Get-PnPListItem -List $listsGet[$listChoice-1].Title -PageSize 500
        ForEach($item in $listItems) {
            $checkItemPermissions = Get-PnPProperty -ClientObject $item -Property "HasUniqueRoleAssignments"
            If($checkItemPermissions) {
                $checkRoleAssignments = Get-PnPProperty -ClientObject $Item -Property RoleAssignments
                ForEach($role in $checkRoleAssignments) {
                    $getMembers = Get-PnPProperty -ClientObject $role -Property RoleDefinitionBindings, Member
                    $getUsers = Get-PnPProperty -ClientObject $role -Property Users -ErrorAction SilentlyContinue
                    
                    If ($role.Member.Title -like "SharingLinks*") {
                        If ($getUsers -ne $null) {
                            ForEach ($user in $getUsers) {
                                $results = New-Object PSObject -Property @{
                                    ListName = $listsGet[$listChoice-1].Title
                                    FileName = $item.FieldValues["FileLeafRef"]
                                    FileType = $item.FieldValues["File_x0020_Type"]
                                    RelativeURL = $item.FieldValues["FileRef"]
                                    User = $user.Title
                                    Access = $role.RoleDefinitionBindings.Name
                                }
                            }
                        }
                    }
                }
                
                reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results
            }
        }

        Disconnect-PnPOnline
        write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
        write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
    } else {
        write-host "`nNO LISTS FOUND." -ForegroundColor Red
    }
}
#endregion

#region DOCUMENT TOOLS FUNCTIONS OPTION "3"
function spoDocumentDeleteAllSharedLinks {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = read-host "`nENTER SITE URL THAT LIST RESIDES ON"
    $results = @()
    
    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue

    $listsGet = @()
    Get-PnPList | where-object { $_.Hidden -eq $false -and $_.BaseTemplate -eq 101 } | foreach-object { $listsGet += ($_) }

    if ($listsGet.count) {
        do {
            write-host "`nPLEASE SELECT A LIST`n"
            foreach ($list in $listsGet) {
                write-host "`t$($listsGet.IndexOf($list)+1): PRESS $($listsGet.IndexOf($list)+1) for $($list.Title)"
            }
            $listChoice = read-host "`nPLEASE MAKE A SELECTION"
        } while (-not($listsGet[$listChoice-1]))

        $listItems = Get-PnPListItem -List $listsGet[$listChoice-1].Title -PageSize 500
        $Batch = New-PnPBatch
        ForEach($item in $listItems) {
            $itemPermission = Get-PnPListItemPermission -List $listsGet[$listChoice-1].Title -Identity $item.Id
            $itemPermission.Permissions | Where {$_.PrincipalName.StartsWith("SharingLinks")} | ForEach-Object {
                $item.RoleAssignments.GetByPrincipalId($_.PrincipalId).DeleteObject()
                Invoke-PnPQuery

                $results = New-Object PSObject -Property @{
                    ListName = $listsGet[$listChoice-1].Title
                    ItemID = $item.Id
                    PrincipalID = $_.PrincipalId
                }
            }
            reportCreate -reportPath "$($setupReportPath)\$($reportName)" -reportData $results
        }
        Invoke-PnPBatch -Batch $Batch

        Disconnect-PnPOnline
        write-host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; write-host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
        write-host "Report Saved: " -ForegroundColor DarkYellow -nonewline; write-host "$($reportPath)\$($reportName)" -ForegroundColor White;
    } else {
        write-host "`nNO LISTS FOUND." -ForegroundColor Red
    }
}
#endregion
#endregion

#region MAIN
showSetup -configFilePath $configFilePath -currentVersion $currentVersion -SetupPath $setupPath -ReportPath $setupReportPath -DirtyWordsPath $setupDirtyWordsPath -DirtyWordsFilePath $setupDirtyWordsFilePath
do {
    showMenu
    $menuMain = read-host "PLEASE MAKE A SELECTION"
    switch -Wildcard ($menuMain) {
        #region SITE TOOLS
        "1*" {
            if ($menuMain.Contains("?")) { 
                write-host "`nCONTAINS TOOLS SPECIFIC TO THE SITE COLLECTION." -ForegroundColor Green
            }

            do {
                showSiteTools
                $menuSub = read-host "PLEASE MAKE A SELECTION"
                switch -Wildcard ($menuSub) {
                    "1*" { 
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES SITE COLLECTION MAP REPORT." -ForegroundColor Green } else {
                            spoSiteMap -reportPath $setupReportPath -reportName "spoSiteMap_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                    "2*" { 
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES SITE COLLECTION PII SCAN REPORT." -ForegroundColor Green } else { 
                            spoSiteScanPII -reportPath $setupReportPath -reportName "spoSiteScanPII_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                    "3*" {
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES SITE COLLECTION ADMIN REPORT." -ForegroundColor Green } else { 
                            spoSiteGetCollectionAdmins -reportPath $setupReportPath -reportName "spoSiteGetCollectionAdmins_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                    "4*" {
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES CAPABILITY TO ADD SITE COLLECTION ADMIN." -ForegroundColor Green } else { 
                            spoSiteAddCollectionAdmin -reportPath $setupReportPath -reportName "spoSiteAddCollectionAdmin_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                    "5*" {
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES CAPABILITY TO DELETE SITE COLLECTION ADMIN." -ForegroundColor Green } else { 
                            spoSiteDeleteCollectionAdmin -reportPath $setupReportPath -reportName "spoSiteDeleteCollectionAdmin_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                    "6*" { 
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES SITE COLLECTION GROUP REPORT." -ForegroundColor Green } else { 
                            spoSiteGetCollectionGroups -reportPath $setupReportPath -reportName "spoSiteGetCollectionGroups_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                }
            } until ($menuSub -eq "e")
        }
        #endregion

        #region USER TOOLS
        "2*" {
            if ($menuMain.Contains("?")) { 
                write-host "`nCONTAINS TOOLS SPECIFIC TO A USER." -ForegroundColor Green
            }

            do {
                showUserTools
                $menuSub = read-host "PLEASE MAKE A SELECTION"
                switch -Wildcard ($menuSub) {
                    "1*" { 
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES CAPABILITY TO DELETE USER FROM SITE COLLECTION (SPO Account will remain)." -ForegroundColor Green } else {
                            spoUserDelete -reportPath $setupReportPath -reportName "spoUserDelete_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                    "2*" { 
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES CAPABILITY TO DELETE ALL USER GROUPS." -ForegroundColor Green } else {
                            spoUserDeleteGroups -reportPath $setupReportPath -reportName "spoUserDeleteGroups_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                }
            } until ($menuSub -eq "e")
        }
        #endregion

        #region LIST TOOLS
        "3*" {
            if ($menuMain.Contains("?")) { 
                write-host "`nCONTAINS TOOLS SPECIFIC TO A LIST" -ForegroundColor Green
            }

            do {
                showListTools
                $menuSub = read-host "PLEASE MAKE A SELECTION"
                switch -wildcard ($menuSub) {
                    "1*" { 
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES CAPABILITY TO SHOW LIST IN BROWSER." -ForegroundColor Green } else {
                            spoListShow -reportPath $setupReportPath -reportName "spoListShow_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                    "2*" { 
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES CAPABILITY TO HIDE LIST FROM BROWSER." -ForegroundColor Green } else {
                            spoListHide -reportPath $setupReportPath -reportName "spoListHide_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                    "3*" { 
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES CAPABILITY TO DELETE ALL UNIQUE PERMISSIONS FROM LIST." -ForegroundColor Green } else {
                            spoListDeleteAllUniquePermissions -reportPath $setupReportPath -reportName "spoListDeleteAllUniquePermissions_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                    "4*" { 
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES CAPABILITY TO DELETE ALL LIST ITEMS." -ForegroundColor Green } else {
                            spoListDeleteAllItems -reportPath $setupReportPath -reportName "spoListDeleteAllItems_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                }
            } until ($menuSub -eq "e")
        }
        #endregion

        #region DOCUMENT TOOLS
        "4*" {
            if ($menuMain.Contains("?")) { 
                write-host "`nCONTAINS TOOLS SPECIFIC TO A DOCUMENT LIBRARY" -ForegroundColor Green
            }

            do {
                showDocumentTools
                $menuSub = read-host "PLEASE MAKE A SELECTION"
                switch ($menuSub) {
                    "1*" { 
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES CAPABILITY TO UPLOAD FOLDER INTO A DOCUMENT LIBRARY." -ForegroundColor Green } else {
                            spoDocumentFolderUpload -reportPath $setupReportPath -reportName "spoDocumentFolderUpload_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                    "2*" { 
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES DOCUMENT LIBRARY SHARED LINK REPORT." -ForegroundColor Green } else {
                            spoDocumentSharedLinks -reportPath $setupReportPath -reportName "spoDocumentSharedLinks_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                    "3*" { 
                        if ($menuSub.Contains("?")) { write-host "`nPROVIDES CAPABILITY TO DELETE ALL DOCUMENT LIBRARY SHARED LINKS." -ForegroundColor Green } else {
                            spoDocumentDeleteAllSharedLinks -reportPath $setupReportPath -reportName "spoDocumentDeleteAllSharedLinks_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                        }
                    }
                }
            } until ($menuSub -eq "e")
        }
        #endregion

        #region SETTINGS
        "s" {
            do {
                showSettings
                $menuSub = read-host "PLEASE MAKE A SELECTION"
                switch ($menuSub) {
                    "1" { start $setupPath }
                    "2" { start $setupDirtyWordsFilePath }
                }
            } until ($menuSub -eq "e")
            showSetup -configFilePath $configFilePath -currentVersion $currentVersion -SetupPath $setupPath -ReportPath $setupReportPath -DirtyWordsPath $setupDirtyWordsPath -DirtyWordsFilePath $setupDirtyWordsFilePath
        }
        #endregion
    }
} until ($menuMain -eq "q")
#endregion
