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

function showMenu {
    Write-Host "
##################################################
#                                                #
#        ░██████╗██████╗░░█████╗░░█████╗░        #
#        ██╔════╝██╔══██╗██╔══██╗██╔══██╗        #
#        ╚█████╗░██████╔╝██║░░██║███████║        #
#        ░╚═══██╗██╔═══╝░██║░░██║██╔══██║        #
#        ██████╔╝██║░░░░░╚█████╔╝██║░░██║        #
#        ╚═════╝░╚═╝░░░░░░╚════╝░╚═╝░░╚═╝        #
#                                                #
#   WELCOME TO THE SHAREPOINT ONLINE ASSISTANT   #
#                                                #
##################################################

##################################################
#                                                #
#                   MAIN MENU                    #
#                                                #
##################################################

##################################################
#                                                #
#        WHICH TOOL WOULD YOU LIKE TO USE?       #
#                                                #
##################################################
    
##################################################
#                                                #
# 1: PRESS '1' FOR PII SCAN.                     #
# 2: PRESS '2' FOR USER GROUP REMOVAL.           #
# S: PRESS 'S' FOR SETTINGS.                     #
# Q: PRESS 'Q' TO QUIT.                          #
#                                                #
##################################################`n"
}

Function Format-FileSize() { # https://community.spiceworks.com/topic/1955251-powershell-help
    Param ([int]$size)
    If ($size -gt 1TB) {[string]::Format("{0:0.00} TB", $size / 1TB)}
    ElseIf ($size -gt 1GB) {[string]::Format("{0:0.00} GB", $size / 1GB)}
    ElseIf ($size -gt 1MB) {[string]::Format("{0:0.00} MB", $size / 1MB)}
    ElseIf ($size -gt 1KB) {[string]::Format("{0:0.00} KB", $size / 1KB)}
    ElseIf ($size -gt 0) {[string]::Format("{0:0.00} B", $size)}
    Else {""}
}

# OPTION "1"
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

                        if (test-path $reportPath) {
                            $results | Select-Object "FileName", "FileExtension", "FileSize", "Path", "Permissions", "Criteria", "Created", "Modified" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation -Append
                        } else {
                            $results | Select-Object "FileName", "FileExtension", "FileSize", "Path", "Permissions", "Criteria", "Created", "Modified" | Export-Csv -Path "$($reportPath)\$($reportName)" -Force -NoTypeInformation
                        }
                    }
                }
            }
        }
    }

    # GET ALL SUB SITE DOCUMENT LIBRARIES
    if ($siteParentOnly -eq $false) {
        $subSites = Get-PnPSubWeb -Recurse # GET ALL SUBSITES

        foreach ($site in $subSites) {
            Connect-PnPOnline -Url $site.Url -UseWebLogin # CONNECT TO SPO SUBSITE
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
                                # Write-Host "$($loginName) - $($rolebindings.Name)" -ForegroundColor Yellow
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

                                if (test-path $reportPath) {
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

# OPTION "S"
function showSettings {   
    Write-Host "
##################################################
#                                                #
#                   SETTINGS                     #
#                                                #
##################################################

##################################################
#                                                #
#             PLEASE SELECT A SETTING            #
#                                                #
##################################################
    
##################################################
#                                                #
# 1: PRESS '1' TO OPEN SPOA FOLDER               #
# 2: PRESS '2' TO OPEN THE DIRTY WORD LIST.      #
# Q: PRESS 'E' TO EXIT BACK TO THE MAIN MENU.    #
#                                                #
##################################################`n"
}

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
        "1" {
            spoScanPII -reportPath $setupReportPath -reportName "SPOSCANPII_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
        }
        "2" {
            spoDeleteUserGroups -reportPath $setupReportPath -reportName "DELETEUSERGROUPS_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
        }
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
        }
    }
} until ($menuMain -eq "q") 
