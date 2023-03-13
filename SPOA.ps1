# MIT License

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

Function Format-FileSize() { # https://community.spiceworks.com/topic/1955251-powershell-help
    Param ([int]$size)
    If ($size -gt 1TB) {[string]::Format("{0:0.00} TB", $size / 1TB)}
    ElseIf ($size -gt 1GB) {[string]::Format("{0:0.00} GB", $size / 1GB)}
    ElseIf ($size -gt 1MB) {[string]::Format("{0:0.00} MB", $size / 1MB)}
    ElseIf ($size -gt 1KB) {[string]::Format("{0:0.00} KB", $size / 1KB)}
    ElseIf ($size -gt 0) {[string]::Format("{0:0.00} B", $size)}
    Else {""}
}

#region SETUP FUNCTIONS
function showSetup {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$SetupPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ReportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$DirtyWordsPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$DirtyWordsFilePath)

    Clear-Host
    $isInstalled=Get-InstalledModule -Name PnP.PowerShell -ErrorAction silentlycontinue
    if($isInstalled.count -eq 0) {
        $Confirm = Read-Host "WOULD YOU LIKE TO INSTALL SHAREPOINT PNP MODULE? [Y] Yes [N] No"

        if($Confirm -match "[yY]") {
            Install-Module -Name PnP.PowerShell -Scope CurrentUser
        } else {
            Write-Host "SharePoint PnP module is needed to perform the functions of this script." -ForegroundColor red
            break
        }
    }

    if (-Not (test-path $SetupPath)) {
        New-Item -Path $SetupPath -ItemType Directory | Out-Null
    }

    if (-Not (test-path $ReportPath)) {
        New-Item -Path $ReportPath -ItemType Directory | Out-Null
    }

    if (-Not (test-path $DirtyWordsPath)) {
        New-Item -Path $DirtyWordsPath -ItemType Directory | Out-Null
    }

    if (-Not (test-path $DirtyWordsFilePath)) {
        $wordDefaultDirtySearchSet = @("\d{3}-\d{3}-\d{4}","\d{3}-\d{2}-\d{4}","MyFitness","CUI","UPMR","SURF","PA","2583","SF86","SF 86","FOUO","GTC","medical","AF469","AF 469","469","Visitor Request","VisitorRequest","Visitor","eQIP","EPR","910","AF910","AF 910","911","AF911","AF 911","OPR","eval","feedback","loc","loa","lor","alpha roster","alpha","roster","recall","SSN","SSAN","AF1466","1466","AF 1466","AF1566","AF 1566","1566","SGLV","SF182","182","SF 182","allocation notice","credit","allocation","2583","AF 1466","AF1466","1466","AF1566","AF 1566","1566","AF469","AF 469","469","AF 422","AF422","422","AF910","AF 910","910","AF911","AF 911","911","AF77","AF 77","77","AF475","AF 475","475","AF707","AF 707","707","AF709","AF 709","709","AF 724","AF724","724","AF912","AF 912","912","AF 931","AF931","931","AF932","AF 932","932","AF948","AF 948","948","AF 3538","AF3538","3538","AF3538E","AF 3538E","AF2096","AF 2096","2096","AF 2098","AF2098","AF 2098","AF 3538","AF3538","3538","1466","1566","469","422","travel","SF128","SF 128","128","SF 86","SF86","86","SGLV","SGLI","DD214","DD 214","214","DD 149","DD149","149") | Select-Object @{Name='Word';Expression={$_}} | Export-Csv $DirtyWordsFilePath -NoType
    }

    if (test-path $DirtyWordsPath) {
        $global:wordDirtySearch = Import-Csv $DirtyWordsFilePath
    }
}
#endregion

#region MAIN AND SETTING MENU FUNCTIONS
function showMenu {
    Write-Host "
###########################################################
#                                                         #
#             ░██████╗██████╗░░█████╗░░█████╗░            #
#             ██╔════╝██╔══██╗██╔══██╗██╔══██╗            #
#             ╚█████╗░██████╔╝██║░░██║███████║            #
#             ░╚═══██╗██╔═══╝░██║░░██║██╔══██║            #
#             ██████╔╝██║░░░░░╚█████╔╝██║░░██║            #
#             ╚═════╝░╚═╝░░░░░░╚════╝░╚═╝░░╚═╝            #
#                                                         #
#        WELCOME TO THE SHAREPOINT ONLINE ASSISTANT       #
#                                                         #
###########################################################`n
###########################################################
#                                                         #
#                        MAIN MENU                        #
#                                                         #
#             WHICH TOOL WOULD YOU LIKE TO USE?           #
#                                                         #
###########################################################`n
###########################################################
#                                                         #
# 1: PRESS '1' FOR SITE TOOLS.                            #
# 2: PRESS '2' FOR USER TOOLS.                            #
# 3: PRESS '3' FOR LIST TOOLS.                            #
# PRESS 'S' FOR SETTINGS OR 'Q' TO QUIT                   #
#                                                         #
###########################################################`n"
}

function showSettings {   
    Write-Host "
###########################################################
#                                                         #
#                        SETTINGS                         #
#                                                         #
#                  PLEASE SELECT A SETTING                #
#                                                         #
###########################################################`n
###########################################################
#                                                         #
# 1: PRESS '1' TO OPEN SPOA FOLDER                        #
# 2: PRESS '2' TO OPEN THE DIRTY WORD LIST.               #
# Q: PRESS 'E' TO EXIT BACK TO THE MAIN MENU.             #
#                                                         #
###########################################################`n"
}
#endregion

#region SITE TOOLS FUNCTIONS
function showSiteTools {   
    Write-Host "
###########################################################
#                                                         #
#                       SITE TOOLS                        #
#                                                         #
#                   PLEASE SELECT A TOOL                  #
#                                                         #
###########################################################`n
###########################################################
#                                                         #
# 1: PRESS '1' FOR SITE MAP REPORT.                       #
# 2: PRESS '2' FOR PII SCAN REPORT.                       #
# 3: PRESS '3' FOR SITE COLLECTION ADMIN REPORT.          #
# 4: PRESS '4' FOR SITE COLLECTION ADMIN DELETE.          #
# 5: PRESS '5' FOR SITE COLLECTION GROUP REPORT.          #
# E: PRESS 'E' TO EXIT BACK TO THE MAIN MENU.             #
#                                                         #
###########################################################`n"
}

# OPTION "1"
function spoSiteMap {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = Read-Host "ENTER SITE COLLECTION URL"
    $results = @()

    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue

    $siteInfo = Get-PnPWeb -Includes Created | select Title, ServerRelativeUrl, Url, Created, Description
    $siteLists = Get-PnPList | Where-Object {$_.Hidden -eq $false}
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
        ListCount = $siteListCount.Count
        ServerRelativeUrl = $siteInfo.ServerRelativeUrl
        Description = $siteInfo.Description
        Created = $siteInfo.Created
    }

    if (test-path "$($reportPath)\$($reportName)") {
        $results | Select-Object "Title", "ServerRelativeUrl", "ListCount", "ItemCount", "Created", "Description" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation -Append
    } else {
        $results | Select-Object "Title", "ServerRelativeUrl", "ListCount", "ItemCount", "Created", "Description" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation
    }

    foreach ($site in $subSites) {
        Connect-PnPOnline -Url $site.Url -UseWebLogin -WarningAction SilentlyContinue
        $subSiteLists = Get-PnPList | Where-Object {$_.Hidden -eq $false}

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
            ListCount = $subSiteListCount.Count
            ItemCount = $subSiteItemCount
            ServerRelativeUrl = $site.ServerRelativeUrl
            Description = $site.Description
            Created = $site.Created
        }

        if (test-path "$($reportPath)\$($reportName)") {
            $results | Select-Object "Title", "ServerRelativeUrl", "ListCount", "ItemCount", "Created", "Description" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation -Append
        } else {
            $results | Select-Object "Title", "ServerRelativeUrl", "ListCount", "ItemCount", "Created", "Description" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation
        }
    }

    # GET TOTAL COUNTS
    $results = New-Object PSObject -Property @{
        Title = "Total"
        ListCount = $siteListCount.Count
        ItemCount = $siteItemCount
        ServerRelativeUrl = $subSites.Count + 1
        Description = ""
        Created = ""
    }

    if (test-path "$($reportPath)\$($reportName)") {
        $results | Select-Object "Title", "ServerRelativeUrl", "ListCount", "ItemCount", "Created", "Description" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation -Append
    } else {
        $results | Select-Object "Title", "ServerRelativeUrl", "ListCount", "ItemCount", "Created", "Description" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation
    }
    Disconnect-PnPOnline

    Write-Host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; Write-Host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    Write-Host "Report Saved: " -ForegroundColor DarkYellow -nonewline; Write-Host "$($reportPath)\$($reportName)" -ForegroundColor White;
}

# OPTION "2"
function spoScanPII {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = Read-Host "ENTER SITE COLLECTION URL"
    $siteParentOnly = $null 
    $results = @()

    $Confirm = Read-Host "WOULD YOU LIKE TO SCAN ALL SUB-SITES? [Y] Yes [N] No"
    if($Confirm -match "[yY]") {
        $siteParentOnly = $false
    } else {
        $siteParentOnly = $true
    }

    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue
    $getDocLibs = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }

    Write-Host "Searching: $($sitePath)" -ForegroundColor Green

    foreach ($DocLib in $getDocLibs) {
        Get-PnPListItem -List $DocLib -Fields "FileRef", "File_x0020_Type", "FileLeafRef", "File_x0020_Size", "Created", "Modified" -PageSize 1000 | Where { $_["FileLeafRef"] -like "*.*" } | Foreach-Object {
            foreach ($word in $global:wordDirtySearch) {
                $wordSearch = "(?i)\b$($word.Word)\b"

                if (($_["FileLeafRef"] -match $wordSearch)) {
                    Write-Host "File found. " -ForegroundColor Red -nonewline; Write-Host "Under: '$($word.Word)' Path: $($_["FileRef"])" -ForegroundColor Yellow;

                    $permissions = @()
                    $perm = Get-PnPProperty -ClientObject $_ -Property RoleAssignments       
                    foreach ($role in $_.RoleAssignments) {
                        $loginName = Get-PnPProperty -ClientObject $role.Member -Property LoginName
                        $rolebindings = Get-PnPProperty -ClientObject $role -Property RoleDefinitionBindings
                        $permissions += "$($loginName) - $($rolebindings.Name)"
                    }
                    $permissions = $permissions | Out-String

                    if ($_ -eq $null) {
                        Write-Host "Error: 'Unable to pull file information'."
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

                        if (test-path "$($reportPath)\$($reportName)") {
                            $results | Select-Object "FileName", "FileExtension", "FileSize", "Path", "Permissions", "Criteria", "Created", "Modified" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation -Append
                        } else {
                            $results | Select-Object "FileName", "FileExtension", "FileSize", "Path", "Permissions", "Criteria", "Created", "Modified" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation
                        }
                    }
                }
            }
        }
    }

    if ($siteParentOnly -eq $false) {
        $subSites = Get-PnPSubWeb -Recurse

        foreach ($site in $subSites) {
            Connect-PnPOnline -Url $site.Url -UseWebLogin -WarningAction SilentlyContinue
            $getSubDocLibs = Get-PnPList | Where-Object {$_.BaseTemplate -eq 101}

            Write-Host "Searching: $($site.Url)" -ForegroundColor Green

            foreach ($subDocLib in $getSubDocLibs) {
                Get-PnPListItem -List $subDocLib -Fields "FileRef", "File_x0020_Type", "FileLeafRef", "File_x0020_Size", "Created", "Modified" -PageSize 1000 | Where { $_["FileLeafRef"] -like "*.*" } | Foreach-Object {
                    foreach ($word in $global:wordDirtySearch) {
                        $wordSearch = "(?i)\b$($word.Word)\b"

                        if (($_["FileLeafRef"] -match $wordSearch)) {
                            Write-Host "File found. " -ForegroundColor Red -nonewline; Write-Host "Under: '$($word.Word)' Path: $($_["FileRef"])" -ForegroundColor Yellow;

                            $permissions = @()
                            $perm = Get-PnPProperty -ClientObject $_ -Property RoleAssignments       
                            foreach ($role in $_.RoleAssignments) {
                                $loginName = Get-PnPProperty -ClientObject $role.Member -Property LoginName
                                $rolebindings = Get-PnPProperty -ClientObject $role -Property RoleDefinitionBindings
                                $permissions += "$($loginName) - $($rolebindings.Name)" 
                            }
                            $permissions = $permissions | Out-String

                            if ($_ -eq $null) {
                                Write-Host "Error: 'Unable to pull file information'."
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

                                if (test-path "$($reportPath)\$($reportName)") {
                                    $results | Select-Object "FileName", "FileExtension", "FileSize", "Path", "Permissions", "Criteria", "Created", "Modified" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation -Append
                                } else {
                                    $results | Select-Object "FileName", "FileExtension", "FileSize", "Path", "Permissions", "Criteria", "Created", "Modified" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    Disconnect-PnPOnline

    Write-Host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; Write-Host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    Write-Host "Report Saved: " -ForegroundColor DarkYellow -nonewline; Write-Host "$($reportPath)\$($reportName)" -ForegroundColor White;
}

# OPTION "3"
function spoGetSiteCollectionAdmins {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = Read-Host "ENTER SITE COLLECTION URL"
    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue

    Get-PnPSiteCollectionAdmin | Select-Object "Id", "Title", "Email", "LoginName", "IsSiteAdmin", "IsShareByEmailGuestUser", "IsHiddenInUI", "PrincipalType" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation

    Disconnect-PnPOnline

    Write-Host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; Write-Host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    Write-Host "Report Saved: " -ForegroundColor DarkYellow -nonewline; Write-Host "$($reportPath)\$($reportName)" -ForegroundColor White;
}

# OPTION "4"
function spoDeleteSiteCollectionAdmin {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = Read-Host "ENTER SITE COLLECTION URL"
    $results = @()
    
    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue
    $getAdmins = @()

    Get-PnPSiteCollectionAdmin | ForEach-Object { $getAdmins += $_.Title }

    do {
        Write-Host "
###########################################################
#                                                         #
#                  PLEASE SELECT AN ADMIN                 #
#                                                         #
###########################################################`n"
        foreach ($admin in $getAdmins) {
            Write-Host "$($getAdmins.IndexOf($admin)+1): PRESS $($getAdmins.IndexOf($admin)+1) for $($admin)"
        }
        $adminChoice = Read-Host "PLEASE MAKE A SELECTION"
    } while (-not($getAdmins[$adminChoice-1]))

    Remove-PnPSiteCollectionAdmin -Owners $getAdmins[$adminChoice-1]

    $results = New-Object PSObject -Property @{
        AdminDeleted = $getAdmins[$adminChoice-1]
    }

    if (-not(test-path "$($reportPath)\$($reportName)")) {
        $results | Select-Object "AdminDeleted" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation
    }
    Disconnect-PnPOnline

    Write-Host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; Write-Host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    Write-Host "Report Saved: " -ForegroundColor DarkYellow -nonewline; Write-Host "$($reportPath)\$($reportName)" -ForegroundColor White;
}

# OPTION "5"
function spoGetSiteCollectionGroups {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = Read-Host "ENTER SITE COLLECTION URL"
    $results = @()

    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue
    Get-PnPGroup | Where {$_.IsHiddenInUI -eq $false -and $_.LoginName -notlike "Limited Access*" -and $_.LoginName -notlike "SharingLinks*"} | Select-Object "Id", "Title", "LoginName", "OwnerTitle" | Foreach-Object {
        $members = @()
        Get-PnPGroupMember -Identity $_.Title | Foreach-Object {
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

        if (test-path "$($reportPath)\$($reportName)") {
            $results | Select-Object "ID", "GroupName", "LoginName", "OwnerTitle", "Members" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation -Append
        } else {
            $results | Select-Object "ID", "GroupName", "LoginName", "OwnerTitle", "Members" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation
        }
    }

    Disconnect-PnPOnline

    Write-Host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; Write-Host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    Write-Host "Report Saved: " -ForegroundColor DarkYellow -nonewline; Write-Host "$($reportPath)\$($reportName)" -ForegroundColor White;
}
#endregion

#region USER TOOLS FUNCTIONS
function showUserTools {   
    Write-Host "
###########################################################
#                                                         #
#                       USER TOOLS                        #
#                                                         #
#                   PLEASE SELECT A TOOL                  #
#                                                         #
###########################################################`n
###########################################################
#                                                         #
# 1: PRESS '1' FOR USER DELETION.                         #
# 1: PRESS '2' FOR ALL ASSIGNED USER GROUP DELETION.      #
# Q: PRESS 'E' TO EXIT BACK TO THE MAIN MENU.             #
#                                                         #
###########################################################`n"
}

# OPTION "1"
function spoDeleteUser {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = Read-Host "ENTER SITE COLLECTION URL"
    $userEmail = Read-Host "ENTER USERS EMAIL"
    $results = @()

    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue
    $userInformation = Get-PnPUser | ? Email -eq $userEmail | ForEach-Object { 
        Write-Host "User Deleted: $($_.Title)" -ForegroundColor Yellow

        Remove-PnPUser -Identity $_.Title -Force

        $results = New-Object PSObject -Property @{
            UserDeleted = $_.Title
            UserEmail = $_.Email
        }

        if (test-path "$($reportPath)\$($reportName)") {
            $results | Select-Object "UserDeleted", "UserEmail" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation -Append
        } else {
            $results | Select-Object "UserDeleted", "UserEmail" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation
        }
    
    
    }
    Disconnect-PnPOnline

    Write-Host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; Write-Host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    Write-Host "Report Saved: " -ForegroundColor DarkYellow -nonewline; Write-Host "$($reportPath)\$($reportName)" -ForegroundColor White;
}

# OPTION "2"
function spoDeleteUserGroups {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = Read-Host "ENTER SITE COLLECTION URL"
    $userEmail = Read-Host "ENTER USERS EMAIL"
    $results = @()

    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue
    $userInformation = Get-PnPUser | ? Email -eq $userEmail | ForEach-Object { $_.Title }
    $userGroups = Get-PnPUser | ? Email -eq $userEmail | Select -ExpandProperty Groups | Where { ($_.Title -notmatch "Limited Access*") -and ($_.Title -notmatch "SharingLinks*") } | ForEach-Object { 
        Write-Host "Name: $userInformation | Group Removed: " -ForegroundColor Yellow -NoNewline; Write-Host $($_.Title) -ForegroundColor Cyan

        Remove-PnPGroupMember -LoginName $userEmail -Identity $_.Title 

        $results = New-Object PSObject -Property @{
            UserDisplay = $userInformation
            UserGroup = $_.Title
        }

        if (test-path "$($reportPath)\$($reportName)") {
            $results | Select-Object "UserDisplay", "UserGroup" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation -Append
        } else {
            $results | Select-Object "UserDisplay", "UserGroup" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation
        }
    }
    Disconnect-PnPOnline

    Write-Host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; Write-Host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    Write-Host "Report Saved: " -ForegroundColor DarkYellow -nonewline; Write-Host "$($reportPath)\$($reportName)" -ForegroundColor White;
}
#endregion

#region LIST TOOLS FUNCTIONS
function showListTools {   
    Write-Host "
###########################################################
#                                                         #
#                       LIST TOOLS                        #
#                                                         #
#                  PLEASE SELECT A SETTING                #
#                                                         #
###########################################################`n
###########################################################
#                                                         #
# 1: PRESS '1' DELETE ALL LIST ITEMS.                     #
# Q: PRESS 'E' TO EXIT BACK TO THE MAIN MENU.             #
#                                                         #
###########################################################`n"
}

# OPTION "1"
function spoDeleteAllListItems {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportPath,
          [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$reportName)

    $sitePath = Read-Host "ENTER SITE URL THAT LIST RESIDES ON"
    $results = @()
    

    Connect-PnPOnline -Url $sitePath -UseWebLogin -WarningAction SilentlyContinue
    $listsGet = @()

    Get-PnPList | Where-Object { $_.Hidden -eq $false -and $_.BaseTemplate -eq 100 } | ForEach-Object { $listsGet += ($_.Title) }

    do {
        Write-Host "
###########################################################
#                                                         #
#                   PLEASE SELECT A LIST                  #
#                                                         #
###########################################################`n"
        foreach ($list in $listsGet) {
            Write-Host "$($listsGet.IndexOf($list)+1): PRESS $($listsGet.IndexOf($list)+1) for $($list)"
        }
        $listChoice = Read-Host "PLEASE MAKE A SELECTION"
    } while (-not($listsGet[$listChoice-1]))

    $listItems =  Get-PnPListItem -List $listsGet[$listChoice-1] -PageSize 500
    $Batch = New-PnPBatch
    ForEach($item in $listItems) {    
         Remove-PnPListItem -List $listsGet[$listChoice-1] -Identity $item.Id -Recycle -Batch $Batch

         $results = New-Object PSObject -Property @{
            ListName = $listsGet[$listChoice-1]
            ItemDeletedID = $item.Id
        }

        if (test-path "$($reportPath)\$($reportName)") {
            $results | Select-Object "ListName", "ItemDeletedID" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation -Append
        } else {
            $results | Select-Object "ListName", "ItemDeletedID" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation
        }
    }
    Invoke-PnPBatch -Batch $Batch

    Disconnect-PnPOnline

    Write-Host "`nCompleted: " -ForegroundColor DarkYellow -nonewline; Write-Host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
    Write-Host "Report Saved: " -ForegroundColor DarkYellow -nonewline; Write-Host "$($reportPath)\$($reportName)" -ForegroundColor White;
}
#endregion

#region MAIN
$setupPath = "C:\users\$env:USERNAME\Documents\SOPA"
$setupReportPath = $setupPath + "\Reports"
$setupDirtyWordsPath = $setupPath + "\DirtyWords"
$setupDirtyWordsFilePath = $setupDirtyWordsPath + "\DirtyWords.csv"

$global:wordDirtySearch = $null;

showSetup -SetupPath $setupPath -ReportPath $setupReportPath -DirtyWordsPath $setupDirtyWordsPath -DirtyWordsFilePath $setupDirtyWordsFilePath
do {
    showMenu
    $menuMain = Read-Host "PLEASE MAKE A SELECTION"
    switch ($menuMain) {
        #region SITE TOOLS
        "1" {
            do {
                showSiteTools
                $menuSiteTools = Read-Host "PLEASE MAKE A SELECTION"
                switch ($menuSiteTools) {
                    "1" {
                        spoSiteMap -reportPath $setupReportPath -reportName "SPOSITEMAP_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                    }
                    "2" {
                        spoScanPII -reportPath $setupReportPath -reportName "SPOSCANPII_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                    }
                    "3" {
                        spoGetSiteCollectionAdmins -reportPath $setupReportPath -reportName "SPOSITECOLLECTIONADMINS_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                    }
                    "4" {
                        spoDeleteSiteCollectionAdmin -reportPath $setupReportPath -reportName "SPOSITECOLLECTIONADMINDELETE_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                    }
                    "5" {
                        spoGetSiteCollectionGroups -reportPath $setupReportPath -reportName "SPOSITECOLLECTIONGROUPS_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                    }
                }
            } until ($menuSiteTools -eq "e")
        }
        #endregion

        #region USER TOOLS
        "2" {
            do {
                showUserTools
                $menuUserTools = Read-Host "PLEASE MAKE A SELECTION"
                switch ($menuUserTools) {
                    "1" {
                        spoDeleteUser -reportPath $setupReportPath -reportName "DELETEUSER_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                    }
                    "2" {
                        spoDeleteUserGroups -reportPath $setupReportPath -reportName "DELETEUSERGROUPS_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                    }
                }
            } until ($menuUserTools -eq "e")
        }
        #endregion

        #region LIST TOOLS
        "3" {
            do {
                showListTools
                $menuUserTools = Read-Host "PLEASE MAKE A SELECTION"
                switch ($menuUserTools) {
                    "1" {
                        spoDeleteAllListItems -reportPath $setupReportPath -reportName "DELETEDLISTITEMS_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
                    }
                }
            } until ($menuUserTools -eq "e")
        }
        #endregion

        #region SETTINGS
        "s" {
            do {
                showSettings
                $menuSettings = Read-Host "PLEASE MAKE A SELECTION"
                switch ($menuSettings) {
                    "1" {
                        start $setupPath
                    }
                    "2" {
                        start $setupDirtyWordsFilePath
                    }
                }
            } until ($menuSettings -eq "e")
            showSetup -SetupPath $setupPath -ReportPath $setupReportPath -DirtyWordsPath $setupDirtyWordsPath -DirtyWordsFilePath $setupDirtyWordsFilePath
        }
        #endregion
    }
} until ($menuMain -eq "q")
#endregion
