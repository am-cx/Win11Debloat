#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess)]
param (
    [switch]$Silent,
    [switch]$Sysprep,
    [string]$LogPath,
    [string]$User,
    [switch]$CreateRestorePoint,
    [switch]$RunAppsListGenerator, [switch]$RunAppConfigurator,
    [switch]$RunDefaults,
    [switch]$RunDefaultsLite,
    [switch]$RunSavedSettings,
    [switch]$RemoveApps, 
    [switch]$RemoveAppsCustom,
    [switch]$RemoveGamingApps,
    [switch]$RemoveCommApps,
    [switch]$RemoveHPApps,
    [switch]$RemoveW11Outlook,
    [switch]$ForceRemoveEdge,
    [switch]$DisableDVR,
    [switch]$DisableGameBarIntegration,
    [switch]$DisableTelemetry,
    [switch]$DisableFastStartup,
    [switch]$DisableModernStandbyNetworking,
    [switch]$DisableBingSearches, [switch]$DisableBing,
    [switch]$DisableDesktopSpotlight,
    [switch]$DisableLockscrTips, [switch]$DisableLockscreenTips,
    [switch]$DisableWindowsSuggestions, [switch]$DisableSuggestions,
    [switch]$DisableEdgeAds,
    [switch]$DisableSettings365Ads,
    [switch]$DisableSettingsHome,
    [switch]$ShowHiddenFolders,
    [switch]$ShowKnownFileExt,
    [switch]$HideDupliDrive,
    [switch]$EnableDarkMode,
    [switch]$DisableTransparency,
    [switch]$DisableAnimations,
    [switch]$TaskbarAlignLeft,
    [switch]$CombineTaskbarAlways, [switch]$CombineTaskbarWhenFull, [switch]$CombineTaskbarNever,
    [switch]$CombineMMTaskbarAlways, [switch]$CombineMMTaskbarWhenFull, [switch]$CombineMMTaskbarNever,
    [switch]$MMTaskbarModeAll, [switch]$MMTaskbarModeMainActive, [switch]$MMTaskbarModeActive,
    [switch]$HideSearchTb, [switch]$ShowSearchIconTb, [switch]$ShowSearchLabelTb, [switch]$ShowSearchBoxTb,
    [switch]$HideTaskview,
    [switch]$DisableStartRecommended,
    [switch]$DisableStartPhoneLink,
    [switch]$DisableCopilot,
    [switch]$DisableRecall,
    [switch]$DisableClickToDo,
    [switch]$DisablePaintAI,
    [switch]$DisableNotepadAI,
    [switch]$DisableEdgeAI,
    [switch]$DisableWidgets, [switch]$HideWidgets,
    [switch]$DisableChat, [switch]$HideChat,
    [switch]$EnableEndTask,
    [switch]$EnableLastActiveClick,
    [switch]$ClearStart,
    [string]$ReplaceStart,
    [switch]$ClearStartAllUsers,
    [string]$ReplaceStartAllUsers,
    [switch]$RevertContextMenu,
    [switch]$DisableMouseAcceleration,
    [switch]$DisableStickyKeys,
    [switch]$HideHome,
    [switch]$HideGallery,
    [switch]$ExplorerToHome,
    [switch]$ExplorerToThisPC,
    [switch]$ExplorerToDownloads,
    [switch]$ExplorerToOneDrive,
    [switch]$NoRestartExplorer,
    [switch]$DisableOnedrive, [switch]$HideOnedrive,
    [switch]$Disable3dObjects, [switch]$Hide3dObjects,
    [switch]$DisableMusic, [switch]$HideMusic,
    [switch]$DisableIncludeInLibrary, [switch]$HideIncludeInLibrary,
    [switch]$DisableGiveAccessTo, [switch]$HideGiveAccessTo,
    [switch]$DisableShare, [switch]$HideShare
)



# Show error if current powershell environment is limited by security policies
if ($ExecutionContext.SessionState.LanguageMode -ne "FullLanguage") {
    Write-Host "错误: Win11Debloat 无法在您的系统上运行, powershell 执行受到安全政策的限制" -ForegroundColor Red
    AwaitKeyToExit
}

# 在指定路径上将脚本输出到 'Win11Debloat.log' 文件夹
if ($LogPath -and (Test-Path $LogPath)) {
    Start-Transcript -Path "$LogPath/Win11Debloat.log" -Append -IncludeInvocationHeader -Force | Out-Null
}
else {
    Start-Transcript -Path "$PSScriptRoot/Win11Debloat.log" -Append -IncludeInvocationHeader -Force | Out-Null
}



##################################################################################################################
#                                                                                                                #
#                                              FUNCTION DEFINITIONS                                              #
#                                                                                                                #
##################################################################################################################



# 显示应用程序选择表单，允许用户选择他们想要删除或保留的应用程序
function ShowAppSelectionForm {
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

    # 初始化表单对象
    $form = New-Object System.Windows.Forms.Form
    $label = New-Object System.Windows.Forms.Label
    $button1 = New-Object System.Windows.Forms.Button
    $button2 = New-Object System.Windows.Forms.Button
    $selectionBox = New-Object System.Windows.Forms.CheckedListBox 
    $loadingLabel = New-Object System.Windows.Forms.Label
    $onlyInstalledCheckBox = New-Object System.Windows.Forms.CheckBox
    $checkUncheckCheckBox = New-Object System.Windows.Forms.CheckBox
    $initialFormWindowState = New-Object System.Windows.Forms.FormWindowState

    $script:selectionBoxIndex = -1

    # saveButton 事件处理程序
    $handler_saveButton_Click= 
    {
        if ($selectionBox.CheckedItems -contains "Microsoft.WindowsStore" -and -not $Silent) {
            $warningSelection = [System.Windows.Forms.Messagebox]::Show('您确定要卸载 Microsoft Store 吗? 此应用程序无法轻易重新安装.', '你确定吗?', 'YesNo', 'Warning')
        
            if ($warningSelection -eq 'No') {
                return
            }
        }

        $script:SelectedApps = $selectionBox.CheckedItems

        # 如果所选应用程序不存在,请创建存储文件
        if (-not (Test-Path "$PSScriptRoot/CustomAppsList")) {
            $null = New-Item "$PSScriptRoot/CustomAppsList"
        }

        Set-Content -Path "$PSScriptRoot/CustomAppsList" -Value $script:SelectedApps

        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }

    # cancelButton 事件处理程序
    $handler_cancelButton_Click= 
    {
        $form.Close()
    }

    $selectionBox_SelectedIndexChanged= 
    {
        $script:selectionBoxIndex = $selectionBox.SelectedIndex
    }

    $selectionBox_MouseDown=
    {
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            if ([System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Shift) {
                if ($script:selectionBoxIndex -ne -1) {
                    $topIndex = $script:selectionBoxIndex

                    if ($selectionBox.SelectedIndex -gt $topIndex) {
                        for (($i = ($topIndex)); $i -le $selectionBox.SelectedIndex; $i++) {
                            $selectionBox.SetItemChecked($i, $selectionBox.GetItemChecked($topIndex))
                        }
                    }
                    elseif ($topIndex -gt $selectionBox.SelectedIndex) {
                        for (($i = ($selectionBox.SelectedIndex)); $i -le $topIndex; $i++) {
                            $selectionBox.SetItemChecked($i, $selectionBox.GetItemChecked($topIndex))
                        }
                    }
                }
            }
            elseif ($script:selectionBoxIndex -ne $selectionBox.SelectedIndex) {
                $selectionBox.SetItemChecked($selectionBox.SelectedIndex, -not $selectionBox.GetItemChecked($selectionBox.SelectedIndex))
            }
        }
    }

    $check_All=
    {
        for (($i = 0); $i -lt $selectionBox.Items.Count; $i++) {
            $selectionBox.SetItemChecked($i, $checkUncheckCheckBox.Checked)
        }
    }

    $load_Apps=
    {
        # 更正表单的初始状态,以防止.Net 最大化表单问题
        $form.WindowState = $initialFormWindowState

        # 在再次加载应用程序列表之前,将状态重置为默认状态
        $script:selectionBoxIndex = -1
        $checkUncheckCheckBox.Checked = $False

        # 显示加载指示器
        $loadingLabel.Visible = $true
        $form.Refresh()

        # 在添加任何新项目之前清除选择框
        $selectionBox.Items.Clear()

        # 设置可以找到Appslist的文件路径
        $appsFile = "$PSScriptRoot/Appslist.txt"
        $listOfApps = ""

        if ($onlyInstalledCheckBox.Checked -and ($script:wingetInstalled -eq $true)) {
            # 尝试通过winget获取已安装应用程序的列表,10秒后超时
            $job = Start-Job { return winget list --accept-source-agreements --disable-interactivity }
            $jobDone = $job | Wait-Job -TimeOut 10

            if (-not $jobDone) {
                # 显示错误，脚本无法从 winget 获取应用程序列表
                [System.Windows.MessageBox]::Show('无法通过 winget 加载已安装应用程序的列表，某些应用程序可能不会显示在列表中.', '错误', '正确', '错误')
            }
            else {
                # 将任务的输出 (应用程序) 添加到 $listOfApps
                $listOfApps = Receive-Job -Job $job
            }
        }

        # 浏览应用程序列表，并将项目逐个添加到 selectionBox
        Foreach ($app in (Get-Content -Path $appsFile | Where-Object { $_ -notmatch '^\s*$' -and $_ -notmatch '^#  .*' -and $_ -notmatch '^# -* #' } )) { 
            $appChecked = $true

            # 如果存在,请删除第一个 # ,并将 appChecked 设置为 false
            if ($app.StartsWith('#')) {
                $app = $app.TrimStart("#")
                $appChecked = $false
            }

            # 从应用程序名称中删除任何注释
            if (-not ($app.IndexOf('#') -eq -1)) {
                $app = $app.Substring(0, $app.IndexOf('#'))
            }
            
            # 从 Appname 中删除前导和后导空格以及`*`字符
            $app = $app.Trim()
            $appString = $app.Trim('*')

            # 确保 appString 不是空的
            if ($appString.length -gt 0) {
                if ($onlyInstalledCheckBox.Checked) {
                    # 如果 onlyInstalledCheckBox 被选中,在将其添加到 selectionBox 之前,请检查应用程序是否已安装
                    if (-not ($listOfApps -like ("*$appString*")) -and -not (Get-AppxPackage -Name $app)) {
                        # 应用程序未安装,继续下一个项目
                        continue
                    }
                    if (($appString -eq "Microsoft.Edge") -and -not ($listOfApps -like "* Microsoft.Edge *")) {
                        # 应用程序未安装,继续下一个项目
                        continue
                    }
                }

                # 将应用程序添加到 selectionBox,并设置其已检查状态
                $selectionBox.Items.Add($appString, $appChecked) | Out-Null
            }
        }
        
        # 隐藏加载指示器
        $loadingLabel.Visible = $False

        # 按字母顺序对 selectionBox 进行排序
        $selectionBox.Sorted = $True
    }

    $form.Text = "Win11Debloat 应用程序选择"
    $form.Name = "appSelectionForm"
    $form.DataBindings.DefaultDataSourceUpdateMode = 0
    $form.ClientSize = New-Object System.Drawing.Size(400,502)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $False

    $button1.TabIndex = 4
    $button1.Name = "saveButton"
    $button1.UseVisualStyleBackColor = $True
    $button1.Text = "确定"
    $button1.Location = New-Object System.Drawing.Point(27,472)
    $button1.Size = New-Object System.Drawing.Size(75,23)
    $button1.DataBindings.DefaultDataSourceUpdateMode = 0
    $button1.add_Click($handler_saveButton_Click)

    $form.Controls.Add($button1)

    $button2.TabIndex = 5
    $button2.Name = "cancelButton"
    $button2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $button2.UseVisualStyleBackColor = $True
    $button2.Text = "取消"
    $button2.Location = New-Object System.Drawing.Point(129,472)
    $button2.Size = New-Object System.Drawing.Size(75,23)
    $button2.DataBindings.DefaultDataSourceUpdateMode = 0
    $button2.add_Click($handler_cancelButton_Click)

    $form.Controls.Add($button2)

    $label.Location = New-Object System.Drawing.Point(13,5)
    $label.Size = New-Object System.Drawing.Size(400,14)
    $Label.Font = 'Microsoft Sans Serif,8'
    $label.Text = '选中您希望删除的应用程序,取消选中您希望保留的应用程序'

    $form.Controls.Add($label)

    $loadingLabel.Location = New-Object System.Drawing.Point(16,46)
    $loadingLabel.Size = New-Object System.Drawing.Size(300,418)
    $loadingLabel.Text = '正在加载应用程序...'
    $loadingLabel.BackColor = "White"
    $loadingLabel.Visible = $false

    $form.Controls.Add($loadingLabel)

    $onlyInstalledCheckBox.TabIndex = 6
    $onlyInstalledCheckBox.Location = New-Object System.Drawing.Point(230,474)
    $onlyInstalledCheckBox.Size = New-Object System.Drawing.Size(150,20)
    $onlyInstalledCheckBox.Text = '仅显示已安装的应用程序'
    $onlyInstalledCheckBox.add_CheckedChanged($load_Apps)

    $form.Controls.Add($onlyInstalledCheckBox)

    $checkUncheckCheckBox.TabIndex = 7
    $checkUncheckCheckBox.Location = New-Object System.Drawing.Point(16,22)
    $checkUncheckCheckBox.Size = New-Object System.Drawing.Size(150,20)
    $checkUncheckCheckBox.Text = '全选/全不选'
    $checkUncheckCheckBox.add_CheckedChanged($check_All)

    $form.Controls.Add($checkUncheckCheckBox)

    $selectionBox.FormattingEnabled = $True
    $selectionBox.DataBindings.DefaultDataSourceUpdateMode = 0
    $selectionBox.Name = "selectionBox"
    $selectionBox.Location = New-Object System.Drawing.Point(13,43)
    $selectionBox.Size = New-Object System.Drawing.Size(374,424)
    $selectionBox.TabIndex = 3
    $selectionBox.add_SelectedIndexChanged($selectionBox_SelectedIndexChanged)
    $selectionBox.add_Click($selectionBox_MouseDown)

    $form.Controls.Add($selectionBox)

    # 保存表单的初始状态
    $initialFormWindowState = $form.WindowState

    # 将应用加载到 selectionBox
    $form.add_Load($load_Apps)

    # 表单打开时聚焦 selectionBox
    $form.Add_Shown({$form.Activate(); $selectionBox.Focus()})

    # 显示表单
    return $form.ShowDialog()
}


# 從指定檔案中返回應用程式列表,它修剪應用程式名稱並刪除任何評論
function ReadAppslistFromFile {
    param (
        $appsFilePath
    )

    $appsList = @()

    # 从提供路径的文件中获取应用程序列表,并逐一删除它们
    Foreach ($app in (Get-Content -Path $appsFilePath | Where-Object { $_ -notmatch '^#.*' -and $_ -notmatch '^\s*$' } )) { 
        # 从应用名称中移除所有注释
        if (-not ($app.IndexOf('#') -eq -1)) {
            $app = $app.Substring(0, $app.IndexOf('#'))
        }

        # 移除应用名称之前和之后的任何空格
        $app = $app.Trim()
        
        $appString = $app.Trim('*')
        $appsList += $appString
    }

    return $appsList
}


# 移除函数调用期间指定的所有用户账户和操作系统镜像中的应用.
function RemoveApps {
    param (
        $appslist
    )

    Foreach ($app in $appsList) { 
        Write-Output "正在尝试移除 $app..."

        # 仅使用 winget 删除 OneDrive 和 Edge
        if (($app -eq "Microsoft.OneDrive") -or ($app -eq "Microsoft.Edge")) {
            if ($script:wingetInstalled -eq $false) {
                Write-Host "错误: WinGet 未安装或已过时,无法移除 $app" -ForegroundColor Red
                continue
            }

            $appName = $app -replace '\.', '_'

            # 通过 winget 卸载应用程序,或创建计划任务以稍后卸载它
            if ($script:Params.ContainsKey("使用者")) {
                RegImport "添加要卸载的预定任务 $app for user $(GetUserName)..." "卸载_$($appName).reg"
            }
            elseif ($script:Params.ContainsKey("系统准备")) {
                RegImport "添加要卸载的预定任务 $app 新用户登录后..." "卸载_$($appName).reg"
            }
            else {
                Strip-Progress -ScriptBlock { winget uninstall --accept-source-agreements --disable-interactivity --id $app } | Tee-Object -Variable wingetOutput
            }

            If (($app -eq "Microsoft.Edge") -and (Select-String -InputObject $wingetOutput -Pattern "卸载失败，退出")) {
                Write-Host "无法通过 Winget 卸载 Microsoft Edge" -ForegroundColor Red
                Write-Output ""

                if ($( Read-Host -Prompt "您想强行卸载 Microsoft Edge 吗? 不推荐! (y/n)" ) -eq 'y') {
                    Write-Output ""
                    ForceRemoveEdge
                }
            }

            continue
        }

        # 使用 Remove-AppxPackage 删除所有其他应用程序
        $app = '*' + $app + '*'

        # 为所有现有用户删除已安装的应用程序
        try {
            Get-AppxPackage -Name $app -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction Continue

            if ($DebugPreference -ne "SilentlyContinue") {
                Write-Host "已为所有用户移除 $app" -ForegroundColor DarkGray
            }
        }
        catch {
            if ($DebugPreference -ne "SilentlyContinue") {
                Write-Host "无法为所有用户移除 $app" -ForegroundColor Yellow
                Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
            }
        }

        # 从操作系统映像中删除已配置的应用程序,因此不会为任何新用户安装该应用程序
        try {
            Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -like $app } | ForEach-Object { Remove-ProvisionedAppxPackage -Online -AllUsers -PackageName $_.PackageName }
        }
        catch {
            Write-Host "无法从windows图像中删除 $app" -ForegroundColor Yellow
            Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
        }
    }
            
    Write-Output ""
}


# 使用卸载程序强制删除 Microsoft Edge
function ForceRemoveEdge {
    # Based on work from loadstring1 & ave9858
    Write-Output "> 正在强制卸载 Microsoft Edge..."

    $regView = [Microsoft.Win32.RegistryView]::Registry32
    $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $regView)
    $hklm.CreateSubKey('SOFTWARE\Microsoft\EdgeUpdateDev').SetValue('AllowUninstall', '')

    # 创建存根 (创建此文件可以卸载 Edge)
    $edgeStub = "$env:SystemRoot\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe"
    New-Item $edgeStub -ItemType Directory | Out-Null
    New-Item "$edgeStub\MicrosoftEdge.exe" | Out-Null

    # 移除 edge
    $uninstallRegKey = $hklm.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge')
    if ($null -ne $uninstallRegKey) {
        Write-Output "运行卸载程序..."
        $uninstallString = $uninstallRegKey.GetValue('UninstallString') + ' --force-uninstall'
        Start-Process cmd.exe "/c $uninstallString" -WindowStyle Hidden -Wait

        Write-Output "正在移除残留文件..."

        $edgePaths = @(
            "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Tombstones\Microsoft Edge.lnk",
            "$env:PUBLIC\Desktop\Microsoft Edge.lnk",
            "$env:USERPROFILE\Desktop\Microsoft Edge.lnk",
            "$edgeStub"
        )

        foreach ($path in $edgePaths) {
            if (Test-Path -Path $path) {
                Remove-Item -Path $path -Force -Recurse -ErrorAction SilentlyContinue
                Write-Host "  已移除 $path" -ForegroundColor DarkGray
            }
        }

        Write-Output "正在清理注册表..."

        # 从自动启动中移除 MS Edge
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "MicrosoftEdgeAutoLaunch_A9F6DCE4ABADF4F51CF45CD7129E3C6C" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "Microsoft Edge Update" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" /v "MicrosoftEdgeAutoLaunch_A9F6DCE4ABADF4F51CF45CD7129E3C6C" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" /v "Microsoft Edge Update" /f *>$null

        Write-Output "Microsoft Edge 已卸载"
    }
    else {
        Write-Output ""
        Write-Host "错误: 无法强制卸载 Microsoft Edge,找不到卸载程序" -ForegroundColor Red
    }
    
    Write-Output ""
}


# 执行提供的命令，并从控制台输出中剥离进度旋转器/条
function Strip-Progress {
    param(
        [ScriptBlock]$ScriptBlock
    )

    # 正模式模式与旋转字符和进度条模式相匹配
    $progressPattern = 'Γû[Æê]|^\s+[-\\|/]\s+$'

    # 更正了大小格式的正则表达式模式,确保使用正确的捕获组
    $sizePattern = '(\d+(\.\d{1,2})?)\s+(B|KB|MB|GB|TB|PB) /\s+(\d+(\.\d{1,2})?)\s+(B|KB|MB|GB|TB|PB)'

    & $ScriptBlock 2>&1 | ForEach-Object {
        if ($_ -is [System.Management.Automation.ErrorRecord]) {
            "错误: $($_.Exception.Message)"
        }
        else {
            $line = $_ -replace $progressPattern, '' -replace $sizePattern, ''
            if (-not ([string]::IsNullOrWhiteSpace($line)) -and -not ($line.StartsWith('  '))) {
                $line
            }
        }
    }
}


# 检查这台机器是否支持 S0 现代待机电源状态.如果支持 S0 Modern Standby,则返回true,否则返回false.
function CheckModernStandbySupport {
    $count = 0

    try {
        switch -Regex (powercfg /a) {
            ':' {
                $count += 1
            }

            '(.*S0.{1,}\))' {
                if ($count -eq 1) {
                    return $true
                }
            }
        }
    }
    catch {
        Write-Host "错误: 无法检查 S0 现代待机支持，powercfg 命令失败" -ForegroundColor Red
        Write-Host ""
        Write-Host "按任意键继续..."
        $null = [System.Console]::ReadKey()
        return $true
    }

    return $false
}


# 返回指定用户的目录路径,如果找不到用户路径,则退出脚本
function GetUserDirectory {
    param (
        $userName,
        $fileName = "",
        $exitIfPathNotFound = $true
    )

    try {
        $userDirectoryExists = Test-Path "$env:SystemDrive\Users\$userName"
        $userPath = "$env:SystemDrive\Users\$userName\$fileName"
    
        if ((Test-Path $userPath) -or ($userDirectoryExists -and (-not $exitIfPathNotFound))) {
            return $userPath
        }
    
        $userDirectoryExists = Test-Path ($env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), "\$userName")
        $userPath = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), "\$userName\$fileName"
    
        if ((Test-Path $userPath) -or ($userDirectoryExists -and (-not $exitIfPathNotFound))) {
            return $userPath
        }
    }
    catch {
        Write-Host "错误:尝试查找用户的用户目录路径时出错 $userName. 请确保用户在此系统上存在." -ForegroundColor Red
        AwaitKeyToExit
    }

    Write-Host "错误: 无法找到用户的用户目录路径 $userName" -ForegroundColor Red
    AwaitKeyToExit
}


# 导入和执行 regfile
function RegImport {
    param (
        $message,
        $path
    )

    Write-Output $message

    if ($script:Params.ContainsKey("Sysprep")) {
        $defaultUserPath = GetUserDirectory -userName "Default" -fileName "NTUSER.DAT"
        
        reg load "HKU\Default" $defaultUserPath | Out-Null
        reg import "$PSScriptRoot\Regfiles\Sysprep\$path"
        reg unload "HKU\Default" | Out-Null
    }
    elseif ($script:Params.ContainsKey("User")) {
        $userPath = GetUserDirectory -userName $script:Params.Item("User") -fileName "NTUSER.DAT"
        
        reg load "HKU\Default" $userPath | Out-Null
        reg import "$PSScriptRoot\Regfiles\Sysprep\$path"
        reg unload "HKU\Default" | Out-Null
        
    }
    else {
        reg import "$PSScriptRoot\Regfiles\$path"  
    }

    Write-Output ""
}


# 重新启动 Windows 资源管理器进程
function RestartExplorer {
    if ($script:Params.ContainsKey("Sysprep") -or $script:Params.ContainsKey("User") -or $script:Params.ContainsKey("NoRestartExplorer")) {
        return
    }

    Write-Output "> 重新启动 Windows 资源管理器进程以应用所有更改... (这可能会导致一些闪烁)"

    if ($script:Params.ContainsKey("DisableMouseAcceleration")) {
        Write-Host "警告:增强指针精度设置更改仅在重新启动后生效" -ForegroundColor Yellow
    }

    if ($script:Params.ContainsKey("DisableStickyKeys")) {
        Write-Host "警告: 粘性键设置更改只有在重新启动后才会生效" -ForegroundColor Yellow
    }

    if ($script:Params.ContainsKey("DisableAnimations")) {
        Write-Host "警告: 只有在重启后才会禁用动画" -ForegroundColor Yellow
    }

    # 仅当powershell进程与操作系统架构匹配时，才重新启动.
    # Restarting explorer from a 32bit PowerShell window will fail on a 64bit OS
    if ([Environment]::Is64BitProcess -eq [Environment]::Is64BitOperatingSystem) {
        Stop-Process -processName: Explorer -Force
    }
    else {
        Write-Warning "无法重新启动 Windows 资源管理器进程,请手动重启PC以应用所有更改."
    }
}


# 替换所有用户的开始菜单，当使用默认的开始菜单模板时，这会清除所有固定的应用程序
# Credit: https://lazyadmin.nl/win-11/customize-windows-11-start-menu-layout/
function ReplaceStartMenuForAllUsers {
    param (
        $startMenuTemplate = "$PSScriptRoot/Assets/Start/start2.bin"
    )

    Write-Output "> Removing all pinned apps from the start menu for all users..."

    # 检查模板bin文件是否存在，如果没有，请尽早返回
    if (-not (Test-Path $startMenuTemplate)) {
        Write-Host "Error: Unable to clear start menu, start2.bin file missing from script folder" -ForegroundColor Red
        Write-Output ""
        return
    }

    # 获取所有用户的启动菜单文件的路径
    $userPathString = GetUserDirectory -userName "*" -fileName "AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState"
    $usersStartMenuPaths = get-childitem -path $userPathString

    # 浏览所有用户并替换开始菜单文件
    ForEach ($startMenuPath in $usersStartMenuPaths) {
        ReplaceStartMenu $startMenuTemplate "$($startMenuPath.Fullname)\start2.bin"
    }

    # 也替换默认用户配置文件的开始菜单文件
    $defaultStartMenuPath = GetUserDirectory -userName "Default" -fileName "AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState" -exitIfPathNotFound $false

    # 如果不存在，请创建文件夹
    if (-not (Test-Path $defaultStartMenuPath)) {
        new-item $defaultStartMenuPath -ItemType Directory -Force | Out-Null
        Write-Output "为默认用户配置文件创建了LocalState文件夹"
    }

    # 将模板复制到默认配置文件
    Copy-Item -Path $startMenuTemplate -Destination $defaultStartMenuPath -Force
    Write-Output "替换了默认用户配置文件的开始菜单"
    Write-Output ""
}


# 替换所有用户的开始菜单，当使用默认的开始菜单模板时，这会清除所有固定的应用程序
# Credit: https://lazyadmin.nl/win-11/customize-windows-11-start-menu-layout/
function ReplaceStartMenu {
    param (
        $startMenuTemplate = "$PSScriptRoot/Assets/Start/start2.bin",
        $startMenuBinFile = "$env:LOCALAPPDATA\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start2.bin"
    )

    # 如果指定了用户，请将路径更改为正确的用户
    if ($script:Params.ContainsKey("User")) {
        $startMenuBinFile = GetUserDirectory -userName "$(GetUserName)" -fileName "AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start2.bin" -exitIfPathNotFound $false
    }

    # 检查模板bin文件是否存在，如果没有，请尽早返回
    if (-not (Test-Path $startMenuTemplate)) {
        Write-Host "Error: 无法替换开始菜单，找不到模板文件" -ForegroundColor Red
        return
    }

    if ([IO.Path]::GetExtension($startMenuTemplate) -ne ".bin" ) {
        Write-Host "Error: 无法替换开始菜单，模板文件不是有效的.bin文件" -ForegroundColor Red
        return
    }

    $userName = [regex]::Match($startMenuBinFile, '(?:Users\\)([^\\]+)(?:\\AppData)').Groups[1].Value

    $backupBinFile = $startMenuBinFile + ".bak"

    if (Test-Path $startMenuBinFile) {
        # 备份当前开始菜单文件
        Move-Item -Path $startMenuBinFile -Destination $backupBinFile -Force
    }
    else {
        Write-Host "Warning: 无法找到用户的原始start2.bin文件 $userName. 未为此用户创建备份!" -ForegroundColor Yellow
        New-Item -ItemType File -Path $startMenuBinFile -Force
    }

    # 复制模板文件
    Copy-Item -Path $startMenuTemplate -Destination $startMenuBinFile -Force

    Write-Output "替换了用户的“开始”菜单 $userName"
}


# 将参数添加到脚本中并写入文件
function AddParameter {
    param (
        $parameterName,
        $message,
        $addToFile = $true
    )

    # 如果密钥尚不存在，请添加密钥
    if (-not $script:Params.ContainsKey($parameterName)) {
        $script:Params.Add($parameterName, $true)
    }

    if (-not $addToFile) {
        Write-Output "- $message"
        return
    }

    # 创建或清除存储上次使用设置的文件
    if (-not (Test-Path "$PSScriptRoot/SavedSettings")) {
        $null = New-Item "$PSScriptRoot/SavedSettings"
    }
    elseif ($script:FirstSelection) {
        $null = Clear-Content "$PSScriptRoot/SavedSettings"
    }
    
    $script:FirstSelection = $false

    # 创建条目并将其添加到文件中
    $entry = "$parameterName#- $message"
    Add-Content -Path "$PSScriptRoot/SavedSettings" -Value $entry
}


function PrintHeader {
    param (
        $title
    )

    $fullTitle = " Win11Debloat Script - $title"

    if ($script:Params.ContainsKey("Sysprep")) {
        $fullTitle = "$fullTitle (系统预制模式)"
    }
    else {
        $fullTitle = "$fullTitle (User: $(GetUserName))"
    }

    Clear-Host
    Write-Host "-------------------------------------------------------------------------------------------"
    Write-Host $fullTitle
    Write-Host "-------------------------------------------------------------------------------------------"
}


function PrintFromFile {
    param (
        $path,
        $title,
        $printHeader = $true
    )

    if ($printHeader) {
        Clear-Host

        PrintHeader $title
    }

    # 从文件中获取和打印脚本菜单
    Foreach ($line in (Get-Content -Path $path )) {   
        Write-Host $line
    }
}


function PrintAppsList {
    param (
        $path,
        $printCount = $false
    )

    if (-not (Test-Path $path)) {
        return
    }
    
    $appsList = ReadAppslistFromFile $path

    if ($printCount) {
        Write-Output "- Remove $($appsList.Count) apps:"
    }

    Write-Host $appsList -ForegroundColor DarkGray
}


function AwaitKeyToExit {
    # Suppress prompt if Silent parameter was passed
    if (-not $Silent) {
        Write-Output ""
        Write-Output "Press any key to exit..."
        $null = [System.Console]::ReadKey()
    }

    Stop-Transcript
    Exit
}


function GetUserName {
    if ($script:Params.ContainsKey("User")) { 
        return $script:Params.Item("User") 
    }
    
    return $env:USERNAME
}


function CreateSystemRestorePoint {
    Write-Output "> 尝试创建系统恢复点..."
    
    $SysRestore = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore" -Name "RPSessionInterval"

    if ($SysRestore.RPSessionInterval -eq 0) {
        if ($Silent -or $( Read-Host -Prompt "系统恢复被禁用，您想启用它并创建一个恢复点吗？? (y/n)") -eq 'y') {
            $enableSystemRestoreJob = Start-Job { 
                try {
                    Enable-ComputerRestore -Drive "$env:SystemDrive"
                }
                catch {
                    Write-Host "Error: 无法启用“系统恢复”: $_" -ForegroundColor Red
                    return
                }
            }
    
            $enableSystemRestoreJobDone = $enableSystemRestoreJob | Wait-Job -TimeOut 20

            if (-not $enableSystemRestoreJobDone) {
                Write-Host "Error: 无法启用系统恢复和创建恢复点，操作超时" -ForegroundColor Red
                return
            }
            else {
                Receive-Job $enableSystemRestoreJob
            }
        }
        else {
            Write-Output ""
            return
        }
    }

    $createRestorePointJob = Start-Job { 
        # 查找不到 24 小时的现有恢复点
        try {
            $recentRestorePoints = Get-ComputerRestorePoint | Where-Object { (Get-Date) - [System.Management.ManagementDateTimeConverter]::ToDateTime($_.CreationTime) -le (New-TimeSpan -Hours 24) }
        }
        catch {
            Write-Host "Error: 无法检索现有的恢复点: $_" -ForegroundColor Red
            return
        }
    
        if ($recentRestorePoints.Count -eq 0) {
            try {
                Checkpoint-Computer -Description "Restore point created by Win11Debloat" -RestorePointType "MODIFY_SETTINGS"
                Write-Output "系统恢复点创建成功"
            }
            catch {
                Write-Host "Error: 无法创建恢复点: $_" -ForegroundColor Red
            }
        }
        else {
            Write-Host "最近的恢复点已经存在，没有创建新的恢复点." -ForegroundColor Yellow
        }
    }
    
    $createRestorePointJobDone = $createRestorePointJob | Wait-Job -TimeOut 20

    if (-not $createRestorePointJobDone) {
        Write-Host "Error: 创建系统恢复点失败，操作超时" -ForegroundColor Red
    }
    else {
        Receive-Job $createRestorePointJob
    }

    Write-Output ""
}


function ShowScriptMenuOptions {
    Do { 
        $ModeSelectionMessage = "请选择一个选项 (1/2/3/0)" 

        PrintHeader 'Menu'

        Write-Host "(1) 默认模式：快速应用建议的更改"
        Write-Host "(2) 自定义模式：手动选择要进行的更改"
        Write-Host "(3) 应用移除模式：选择并移除应用，无需进行其他更改"

        # 仅当SavedSettings文件存在时才显示此选项
        if (Test-Path "$PSScriptRoot/SavedSettings") {
            Write-Host "(4) 应用上次保存的自定义设置"
            
            $ModeSelectionMessage = "请选择一个选项 (1/2/3/4/0)" 
        }

        Write-Host ""
        Write-Host "(0) 显示更多信息"
        Write-Host ""
        Write-Host ""

        $Mode = Read-Host $ModeSelectionMessage

        if ($Mode -eq '0') {
            # Print information screen from file
            PrintFromFile "$PSScriptRoot/Assets/Menus/Info" "Information"

            Write-Host "按任意键返回..."
            $null = [System.Console]::ReadKey()
        }
        elseif (($Mode -eq '4') -and -not (Test-Path "$PSScriptRoot/SavedSettings")) {
            $Mode = $null
        }
    }
    while ($Mode -ne '1' -and $Mode -ne '2' -and $Mode -ne '3' -and $Mode -ne '4')

    return $Mode
}


function ShowDefaultMode {
    AddParameter 'CreateRestorePoint' '创建系统恢复点' $false

    # 显示删除应用程序的选项，或者如果通过了RunDefaults或RunDefaultsLite参数，则设置选择
    if ($RunDefaults) {
        $RemoveAppsInput = '1'
    }
    elseif ($RunDefaultsLite) {
        $RemoveAppsInput = '0'                
    }
    else {
        $RemoveAppsInput = ShowDefaultModeAppRemovalOptions

        if (($script:selectedApps.contains('Microsoft.XboxGameOverlay') -or $script:selectedApps.contains('Microsoft.XboxGamingOverlay')) -and 
          $( Read-Host -Prompt "禁用游戏栏集成和游戏/屏幕录制？ 这也阻止了ms-gamingoverlay和ms-gamebar弹出窗口 (y/n)" ) -eq 'y') {
            $DisableGameBarIntegrationInput = $true;
        }
    }

    PrintHeader 'Default Mode'

    Write-Output "Win11Debloat will make the following changes:"

    # 根据用户输入选择正确的选项
    switch ($RemoveAppsInput) {
        '1' {
            AddParameter 'RemoveApps' 'Remove the default selection of apps:' $false
            PrintAppsList "$PSScriptRoot/Appslist.txt"
        }
        '2' {
            AddParameter 'RemoveAppsCustom' "Remove $($script:SelectedApps.Count) apps:" $false
            PrintAppsList "$PSScriptRoot/CustomAppsList"
        }
    }

    if ($DisableGameBarIntegrationInput) {
        AddParameter 'DisableDVR' 'Disable Xbox game/screen recording' $false
        AddParameter 'DisableGameBarIntegration' 'Disable Game Bar integration' $false
    }

    # 仅为Windows 10用户添加此选项
    if (get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'") {
        AddParameter 'Hide3dObjects' "在文件资源管理器中隐藏“此电脑”下的3D对象文件夹" $false
        AddParameter 'HideChat' 'Hide the chat (meet now) icon from the taskbar' $false
    }

    # 仅为Windows 11用户添加这些选项 (build 22000+)
    if ($WinVersion -ge 22000) {
        if ($script:ModernStandbySupported) {
            AddParameter 'DisableModernStandbyNetworking' '在现代待机期间禁用网络连接' $false
        }

        AddParameter 'DisableRecall' 'Disable Windows Recall' $false
        AddParameter 'DisableClickToDo' 'Disable Click to Do (AI text & image analysis)' $false
    }

    PrintFromFile "$PSScriptRoot/Assets/Menus/DefaultSettings" "Default Mode" $false

    # 如果静音参数已通过，请抑制提示
    if (-not $Silent) {
        Write-Output "按回车键执行脚本或按CTRL+C退出..."
        Read-Host | Out-Null
    } 

    $DefaultParameterNames = 'DisableCopilot','DisableTelemetry','DisableSuggestions','DisableEdgeAds','DisableLockscreenTips','DisableBing','ShowKnownFileExt','DisableWidgets','DisableFastStartup'

    PrintHeader 'Default Mode'

    # 添加默认参数，如果它们尚未存在
    foreach ($ParameterName in $DefaultParameterNames) {
        if (-not $script:Params.ContainsKey($ParameterName)) {
            $script:Params.Add($ParameterName, $true)
        }
    }
}


function ShowDefaultModeAppRemovalOptions {
    PrintHeader 'Default Mode'

    Write-Host "请注意：默认选择的应用程序包括微软团队、Spotify、便笺等。 选择选项2来验证并更改脚本删除的应用程序." -ForegroundColor DarkGray
    Write-Host ""

    Do {
        Write-Host "选项:" -ForegroundColor Yellow
        Write-Host " (n) 不要删除任何应用程序" -ForegroundColor Yellow
        Write-Host " (1) 僅移除預設的應用程式選擇" -ForegroundColor Yellow
        Write-Host " (2) 手动选择要删除的应用程序" -ForegroundColor Yellow
        $RemoveAppsInput = Read-Host“您要移除任何应用程序吗？ 应用程序将针对所有用户被删除 (n/1/2)"

        # 如果用户输入了选项，则显示应用程序选择表单 3
        if ($RemoveAppsInput -eq '2') {
            $result = ShowAppSelectionForm

            if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
                # User cancelled or closed app selection, show error and change RemoveAppsInput so the menu will be shown again
                Write-Host ""
                Write-Host "取消申请选择，请重试" -ForegroundColor Red

                $RemoveAppsInput = 'c'
            }
            
            Write-Host ""
        }
    }
    while ($RemoveAppsInput -ne 'n' -and $RemoveAppsInput -ne '0' -and $RemoveAppsInput -ne '1' -and $RemoveAppsInput -ne '2')

    return $RemoveAppsInput
}


function ShowCustomModeOptions {
    # 获取当前的Windows构建版本，以便与功能进行比较
    $WinVersion = Get-ItemPropertyValue 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' CurrentBuild
            
    PrintHeader 'Custom Mode'

    AddParameter 'CreateRestorePoint' 'Create a system restore point'

    # Show options for removing apps, only continue on valid input
    Do {
        Write-Host "选项:" -ForegroundColor Yellow
        Write-Host " (n) 不要删除任何应用程序" -ForegroundColor Yellow
        Write-Host " (1) 僅移除預設的應用程式選擇" -ForegroundColor Yellow
        Write-Host " (2) 删除应用程序的默认选择，以及邮件和日历应用程序以及游戏相关应用程序"  -ForegroundColor Yellow
        Write-Host " (3) 手动选择要删除的应用程序" -ForegroundColor Yellow
        $RemoveAppsInput = Read-Host "您想移除任何应用程序吗？ 应用程序将针对所有用户被删除 (n/1/2/3)"

        # 如果用户输入了选项，则显示应用程序选择表单 3
        if ($RemoveAppsInput -eq '3') {
            $result = ShowAppSelectionForm

            if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
                # User cancelled or closed app selection, show error and change RemoveAppsInput so the menu will be shown again
                Write-Output ""
                Write-Host "取消申请选择，请重试" -ForegroundColor Red

                $RemoveAppsInput = 'c'
            }
            
            Write-Output ""
        }
    }
    while ($RemoveAppsInput -ne 'n' -and $RemoveAppsInput -ne '0' -and $RemoveAppsInput -ne '1' -and $RemoveAppsInput -ne '2' -and $RemoveAppsInput -ne '3') 

    # 根据用户输入选择正确的选项
    switch ($RemoveAppsInput) {
        '1' {
            AddParameter 'RemoveApps' 'Remove the default selection of apps'
        }
        '2' {
            AddParameter 'RemoveApps' '移除應用程式的預設選擇'
            AddParameter 'RemoveCommApps' '移除“邮件”、“日历”和“人员”应用'
            AddParameter 'RemoveW11Outlook' '删除新的Outlook for Windows应用程序'
            AddParameter 'RemoveGamingApps' '删除Xbox应用程序和Xbox游戏栏'
            AddParameter 'DisableDVR' '停用Xbox游戏/屏幕录制'
            AddParameter 'DisableGameBarIntegration' '禁用游戏栏集成'
        }
        '3' {
            Write-Output "You have selected $($script:SelectedApps.Count) apps for removal"

            AddParameter 'RemoveAppsCustom' "Remove $($script:SelectedApps.Count) apps:"

            Write-Output ""

            if ($( Read-Host -Prompt "禁用游戏栏集成和游戏/屏幕录制？ 这也阻止了ms-gamingoverlay和ms-gamebar弹出窗口 (y/n)" ) -eq 'y') {
                AddParameter 'DisableDVR' 'Disable Xbox game/screen recording'
                AddParameter 'DisableGameBarIntegration' '禁用游戏栏集成'
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "禁用遥测、诊断数据、活动历史记录、应用程序启动跟踪和定向广告? (y/n)" ) -eq 'y') {
        AddParameter 'DisableTelemetry' 'Disable telemetry, diagnostic data, activity history, app-launch tracking & targeted ads'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "在开始、设置、通知、资源管理器、锁定屏幕和 Edge? (y/n)" ) -eq 'y') {
        AddParameter 'DisableSuggestions' 'Disable tips, tricks, suggestions and ads in start, settings, notifications and File Explorer'
        AddParameter 'DisableEdgeAds' 'Disable ads, suggestions and the MSN news feed in Microsoft Edge'
        AddParameter 'DisableSettings365Ads' 'Disable Microsoft 365 ads in Settings Home'
        AddParameter 'DisableLockscreenTips' 'Disable tips & tricks on the lockscreen'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "从Windows搜索中禁用并删除Bing网页搜索、Bing AI和Cortana? (y/n)" ) -eq 'y') {
        AddParameter 'DisableBing' '从Windows搜索中禁用并删除Bing网页搜索、Bing AI和Cortana'
    }

    # Only show this option for Windows 11 users running build 22621 or later
    if ($WinVersion -ge 22621) {
        Write-Output ""

        # 显示禁用/删除人工智能功能的选项，仅在有效输入时继续
        Do {
            Write-Host "选项:" -ForegroundColor Yellow
            Write-Host " (n) 不要禁用任何人工智能功能" -ForegroundColor Yellow
            Write-Host " (1) 禁用Microsoft Copilot、Windows召回和点击操作" -ForegroundColor Yellow
            Write-Host " (2) 在Microsoft Edge、Paint和Notepad中禁用Microsoft Copilot、Windows Recall、Click to Do和AI功能"  -ForegroundColor Yellow
            $DisableAIInput = Read-Host "你想禁用任何人工智能功能吗？ 这适用于所有用户 (n/1/2)"
        }
        while ($DisableAIInput -ne 'n' -and $DisableAIInput -ne '0' -and $DisableAIInput -ne '1' -and $DisableAIInput -ne '2') 

        # 根据用户输入选择正确的选项
        switch ($DisableAIInput) {
            '1' {
                AddParameter 'DisableCopilot' '禁用并删除Microsoft Copilot'
                AddParameter 'DisableRecall' '停用Windows召回'
                AddParameter 'DisableClickToDo' '禁用点击做（人工智能文本和图像分析)'
            }
            '2' {
                AddParameter 'DisableCopilot' '禁用并删除Microsoft Copilot'
                AddParameter 'DisableRecall' '停用Windows召回'
                AddParameter 'DisableClickToDo' '禁用点击做（人工智能文本和图像分析)'
                AddParameter 'DisableEdgeAI' '禁用人工智能功能 Edge'
                AddParameter 'DisablePaintAI' '在Paint中禁用AI功能'
                AddParameter 'DisableNotepadAI' '禁用记事本中的人工智能功能'
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "在桌面上禁用Windows聚光灯背景? (y/n)" ) -eq 'y') {
        AddParameter 'DisableDesktopSpotlight' 'Disable the Windows Spotlight desktop background option.'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "为系统和应用程序启用深色模式? (y/n)" ) -eq 'y') {
        AddParameter 'EnableDarkMode' 'Enable dark mode for system and apps'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "禁用透明度、动画和视觉效果? (y/n)" ) -eq 'y') {
        AddParameter 'DisableTransparency' 'Disable transparency effects'
        AddParameter 'DisableAnimations' 'Disable animations and visual effects'
    }

    # Only show this option for Windows 11 users running build 22000 or later
    if ($WinVersion -ge 22000) {
        Write-Output ""

        if ($( Read-Host -Prompt "恢复旧的Windows 10样式上下文菜单? (y/n)" ) -eq 'y') {
            AddParameter 'RevertContextMenu' '恢复旧的Windows 10样式上下文菜单'
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "关闭增强指针精度，也称为鼠标加速? (y/n)" ) -eq 'y') {
        AddParameter 'DisableMouseAcceleration' '关闭增强指针精度（鼠标加速)'
    }

    # Only show this option for Windows 11 users running build 26100 or later
    if ($WinVersion -ge 26100) {
        Write-Output ""

        if ($( Read-Host -Prompt "禁用粘性键键盘快捷键? (y/n)" ) -eq 'y') {
            AddParameter 'DisableStickyKeys' '禁用粘性键键盘快捷键'
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "禁用快速启动？ 这适用于所有用户 (y/n)" ) -eq 'y') {
        AddParameter 'DisableFastStartup' 'Disable Fast Start-up'
    }

    # 仅为运行 build 22000 或更高版本的 Windows 11 用户显示此选项，并且机器至少有一个电池
    if (($WinVersion -ge 22000) -and $script:ModernStandbySupported) {
        Write-Output ""

        if ($( Read-Host -Prompt "在现代待机期间禁用网络连接？ 这适用于所有用户 (y/n)" ) -eq 'y') {
            AddParameter 'DisableModernStandbyNetworking' '在现代待机期间禁用网络连接'
        }
    }

    # 仅显示为Windows 10用户禁用上下文菜单项的选项，或者如果用户选择恢复Windows 10上下文菜单
    if ((get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'") -or $script:Params.ContainsKey('RevertContextMenu')) {
        Write-Output ""

        if ($( Read-Host -Prompt "您想禁用任何上下文菜单选项吗? (y/n)" ) -eq 'y') {
            Write-Output ""

            if ($( Read-Host -Prompt "   在上下文菜单中隐藏“包含在库中”选项? (y/n)" ) -eq 'y') {
                AddParameter 'HideIncludeInLibrary' "在上下文菜单中隐藏“包含在库中”选项"
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   隐藏上下文菜单中的“提供访问权限”选项? (y/n)" ) -eq 'y') {
                AddParameter 'HideGiveAccessTo' "隐藏上下文菜单中的“提供访问权限”选项"
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   在上下文菜单中隐藏“共享”选项? (y/n)" ) -eq 'y') {
                AddParameter 'HideShare' "在上下文菜单中隐藏“共享”选项"
            }
        }
    }

    # Only show this option for Windows 11 users running build 22621 or later
    if ($WinVersion -ge 22621) {
        Write-Output ""

        if ($( Read-Host -Prompt "你想对开始菜单进行任何更改吗? (y/n)" ) -eq 'y') {
            Write-Output ""

            if ($script:Params.ContainsKey("Sysprep")) {
                if ($( Read-Host -Prompt "从所有现有和新用户的开始菜单中删除所有固定的应用程序? (y/n)" ) -eq 'y') {
                    AddParameter 'ClearStartAllUsers' '从现有和新用户的“开始”菜单中删除所有固定的应用程序'
                }
            }
            else {
                Do {
                    Write-Host "   选项:" -ForegroundColor Yellow
                    Write-Host "    (n) 不要从 "开始" 菜单中移除任何固定的应用程序" -ForegroundColor Yellow
                    Write-Host "    (1) 仅针对此用户从开始菜单中删除所有固定的应用程序 ($(GetUserName))" -ForegroundColor Yellow
                    Write-Host "    (2) 从所有现有和新用户的开始菜单中删除所有固定的应用程序"  -ForegroundColor Yellow
                    $ClearStartInput = Read-Host "   从开始菜单中移除所有固定的应用程序? (n/1/2)" 
                }
                while ($ClearStartInput -ne 'n' -and $ClearStartInput -ne '0' -and $ClearStartInput -ne '1' -and $ClearStartInput -ne '2') 

                # 根据用户输入选择正确的选项
                switch ($ClearStartInput) {
                    '1' {
                        AddParameter 'ClearStart' "仅针对此用户从开始菜单中删除所有固定的应用程序"
                    }
                    '2' {
                        AddParameter 'ClearStartAllUsers' "从所有现有和新用户的开始菜单中删除所有固定的应用程序"
                    }
                }
            }

            # 不要为运行版本26200 及以上版本的用户显示选项，因为此版本中已删除此设置
            if ($WinVersion -lt 26200) {
                Write-Output ""

                if ($( Read-Host -Prompt "   禁用 "开始" 菜单中的推荐部分？ 这适用于所有用户 (y/n)" ) -eq 'y') {
                    AddParameter 'DisableStartRecommended' '禁用开始菜单中的推荐部分.'
                }
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   在开始菜单中禁用Phone Link移动设备集成? (y/n)" ) -eq 'y') {
                AddParameter 'DisableStartPhoneLink' '在开始菜单中禁用 Phone Link 移动设备集成.'
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "您想对任务栏和相关服务进行任何更改吗? (y/n)" ) -eq 'y') {
        # Only show these specific options for Windows 11 users running build 22000 or later
        if ($WinVersion -ge 22000) {
            Write-Output ""

            if ($( Read-Host -Prompt "   A左侧的工作列按钮? (y/n)" ) -eq 'y') {
                AddParameter 'TaskbarAlignLeft' 'Align taskbar icons to the left'
            }

            # 在任务栏上显示组合图标的选项，仅在有效输入时继续
            Do {
                Write-Output ""
                Write-Host "   选项:" -ForegroundColor Yellow
                Write-Host "    (n) 不修改" -ForegroundColor Yellow
                Write-Host "    (1) 总是" -ForegroundColor Yellow
                Write-Host "    (2) 当任务栏已满时" -ForegroundColor Yellow
                Write-Host "    (3) 从不" -ForegroundColor Yellow
                $TbCombineTaskbar = Read-Host "   Combine taskbar buttons and hide labels? (n/1/2/3)" 
            }
            while ($TbCombineTaskbar -ne 'n' -and $TbCombineTaskbar -ne '0' -and $TbCombineTaskbar -ne '1' -and $TbCombineTaskbar -ne '2' -and $TbCombineTaskbar -ne '3') 

            # 根据用户输入选择正确的任务栏组选项
            switch ($TbCombineTaskbar) {
                '1' {
                    AddParameter 'CombineTaskbarAlways' '始终合并工作列按钮并隐藏主显示器的标签'
                    AddParameter 'CombineMMTaskbarAlways' '始终合并任务栏按钮并隐藏二级显示器的标签'
                }
                '2' {
                    AddParameter 'CombineTaskbarWhenFull' '当主显示器的工作列已满时，合并任务栏按钮并隐藏标签'
                    AddParameter 'CombineMMTaskbarWhenFull' '当任务栏已满时，将任务栏按钮合并并隐藏二级显示器的标签'
                }
                '3' {
                    AddParameter 'CombineTaskbarNever' '切勿将任务栏按钮组合在一起，并显示主显示器的标签'
                    AddParameter 'CombineMMTaskbarNever' '切勿将任务栏按钮合并为辅助显示器显示标签'
                }
            }

            # Show options for changing on what taskbar(s) app icons are shown, only continue on valid input
            Do {
                Write-Output ""
                Write-Host "   选项:" -ForegroundColor Yellow
                Write-Host "    (n) 不修改" -ForegroundColor Yellow
                Write-Host "    (1) 在所有任务栏上显示应用程序图标" -ForegroundColor Yellow
                Write-Host "    (2) 在主任务栏和窗口打开的工作列上显示应用程序图标" -ForegroundColor Yellow
                Write-Host "    (3) 仅在窗口打开的工作列上显示应用程序图标" -ForegroundColor Yellow
                $TbCombineTaskbar = Read-Host "   Change how to show app icons on the taskbar when using multiple monitors? (n/1/2/3)" 
            }
            while ($TbCombineTaskbar -ne 'n' -and $TbCombineTaskbar -ne '0' -and $TbCombineTaskbar -ne '1' -and $TbCombineTaskbar -ne '2' -and $TbCombineTaskbar -ne '3') 

            # 根据用户输入选择正确的任务栏组选项
            switch ($TbCombineTaskbar) {
                '1' {
                    AddParameter 'MMTaskbarModeAll' '在所有任务栏上显示应用程序图标'
                }
                '2' {
                    AddParameter 'MMTaskbarModeMainActive' '在主任务栏和窗口打开的工作列上显示应用程序图标'
                }
                '3' {
                    AddParameter 'MMTaskbarModeActive' '仅在窗口打开的工作列上显示应用程序图标'
                }
            }

            # 在工作列上显示搜索图标的选项，仅在有效输入时继续
            Do {
                Write-Output ""
                Write-Host "   选项:" -ForegroundColor Yellow
                Write-Host "    (n) 不修改" -ForegroundColor Yellow
                Write-Host "    (1) 从任务栏中隐藏搜索图标" -ForegroundColor Yellow
                Write-Host "    (2) 在任务栏上显示搜索图标" -ForegroundColor Yellow
                Write-Host "    (3) 在工作列上显示带有标签的搜索图标" -ForegroundColor Yellow
                Write-Host "    (4) 在任务栏上显示搜索框" -ForegroundColor Yellow
                $TbSearchInput = Read-Host "   Hide or change the search icon on the taskbar? (n/1/2/3/4)" 
            }
            while ($TbSearchInput -ne 'n' -and $TbSearchInput -ne '0' -and $TbSearchInput -ne '1' -and $TbSearchInput -ne '2' -and $TbSearchInput -ne '3' -and $TbSearchInput -ne '4') 

            # 根据用户输入选择正确的任务栏搜索选项
            switch ($TbSearchInput) {
                '1' {
                    AddParameter '隐藏搜索Tb' '从任务栏中隐藏搜索图标'
                }
                '2' {
                    AddParameter '显示搜索图标Tb' '在任务栏上显示搜索图标'
                }
                '3' {
                    AddParameter '显示搜索标签Tb' '在工作列上显示带有标签的搜索图标'
                }
                '4' {
                    AddParameter '显示搜索框Tb' '在任务栏上显示搜索框'
                }
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   从任务栏中隐藏任务视图按钮? (y/n)" ) -eq 'y') {
                AddParameter 'HideTaskview' 'Hide the taskview button from the taskbar'
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   禁用小部件服务以删除任务栏和锁屏上的小部件? (y/n)" ) -eq 'y') {
            AddParameter 'DisableWidgets' 'Disable widgets on the taskbar & lockscreen'
        }

        # Only show this options for Windows users running build 22621 or earlier
        if ($WinVersion -le 22621) {
            Write-Output ""

            if ($( Read-Host -Prompt "   从任务栏中隐藏聊天(立即见面)图标? (y/n)" ) -eq 'y') {
                AddParameter 'HideChat' 'Hide the chat (meet now) icon from the taskbar'
            }
        }
        
        # Only show this options for Windows users running build 22631 or later
        if ($WinVersion -ge 22631) {
            Write-Output ""

            if ($( Read-Host -Prompt "   在任务栏右键单击菜单中启用 "结束任务" 选项? (y/n)" ) -eq 'y') {
                AddParameter 'EnableEndTask' "Enable the 'End Task' option in the taskbar right click menu"
            }
        }
        
        Write-Output ""
        if ($( Read-Host -Prompt "   在任务栏应用程序区域中启用 "最后活跃点击" 行为? (y/n)" ) -eq 'y') {
            AddParameter 'EnableLastActiveClick' "Enable the 'Last Active Click' behavior in the taskbar app area"
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "您想对文件资源管理器进行任何更改吗? (y/n)" ) -eq 'y') {
        # Show options for changing the File Explorer default location
        Do {
            Write-Output ""
            Write-Host "   Options:" -ForegroundColor Yellow
            Write-Host "    (n) 不改变" -ForegroundColor Yellow
            Write-Host "    (1) 打开文件资源管理器以 'Home'" -ForegroundColor Yellow
            Write-Host "    (2) 打开文件资源管理器以 'This PC'" -ForegroundColor Yellow
            Write-Host "    (3) 打开文件资源管理器以 'Downloads'" -ForegroundColor Yellow
            Write-Host "    (4) 打开文件资源管理器以 'OneDrive'" -ForegroundColor Yellow
            $ExplSearchInput = Read-Host "   更改文件资源管理器打开的默认位置? (n/1/2/3/4)" 
        }
        while ($ExplSearchInput -ne 'n' -and $ExplSearchInput -ne '0' -and $ExplSearchInput -ne '1' -and $ExplSearchInput -ne '2' -and $ExplSearchInput -ne '3' -and $ExplSearchInput -ne '4') 

        # 根据用户输入选择正确的任务栏搜索选项
        switch ($ExplSearchInput) {
            '1' {
                AddParameter 'ExplorerToHome' "更改文件资源管理器打开的默认位置 'Home'"
            }
            '2' {
                AddParameter 'ExplorerToThisPC' "更改文件资源管理器打开的默认位置 '这台电脑'"
            }
            '3' {
                AddParameter 'ExplorerToDownloads' "更改文件资源管理器打开的默认位置 '下载'"
            }
            '4' {
                AddParameter 'ExplorerToOneDrive' "更改文件资源管理器打开的默认位置 'OneDrive'"
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   显示隐藏的文件,文件夹和驱动器? (y/n)" ) -eq 'y') {
            AddParameter 'ShowHiddenFolders' 'Show hidden files, folders and drives'
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   顯示已知檔案型別的副檔名? (y/n)" ) -eq 'y') {
            AddParameter 'ShowKnownFileExt' 'Show file extensions for known file types'
        }

        # Only show this option for Windows 11 users running build 22000 or later
        if ($WinVersion -ge 22000) {
            Write-Output ""

            if ($( Read-Host -Prompt "   从文件资源管理器侧面板中隐藏主页部分? (y/n)" ) -eq 'y') {
                AddParameter 'HideHome' 'Hide the Home section from the File Explorer sidepanel'
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   从文件资源管理器侧面板中隐藏画廊部分? (y/n)" ) -eq 'y') {
                AddParameter 'HideGallery' 'Hide the Gallery section from the File Explorer sidepanel'
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   从文件资源管理器侧板中隐藏重复的可移动驱动器条目，以便它们仅显示在此PC下? (y/n)" ) -eq 'y') {
            AddParameter 'HideDupliDrive' 'Hide duplicate removable drive entries from the File Explorer sidepanel'
        }

        # 仅显示 Windows 10 用户禁用这些特定文件夹的选项
        if (get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'") {
            Write-Output ""

            if ($( Read-Host -Prompt "您想从文件资源管理器侧面板中隐藏任何文件夹吗? (y/n)" ) -eq 'y') {
                Write-Output ""

                if ($( Read-Host -Prompt "   从文件资源管理器侧面板中隐藏OneDrive文件夹? (y/n)" ) -eq 'y') {
                    AddParameter 'HideOnedrive' 'Hide the OneDrive folder in the File Explorer sidepanel'
                }

                Write-Output ""
                
                if ($( Read-Host -Prompt "   从文件资源管理器侧面板中隐藏3D对象文件夹? (y/n)" ) -eq 'y') {
                    AddParameter 'Hide3dObjects' "Hide the 3D objects folder under 'This pc' in File Explorer" 
                }
                
                Write-Output ""

                if ($( Read-Host -Prompt "   从文件资源管理器侧面板中隐藏音乐文件夹? (y/n)" ) -eq 'y') {
                    AddParameter 'HideMusic' "Hide the music folder under 'This pc' in File Explorer"
                }
            }
        }
    }

    # 如果静音参数已通过，请抑制提示
    if (-not $Silent) {
        Write-Output ""
        Write-Output ""
        Write-Output ""
        Write-Output "按回车键确认您的选择并执行脚本，或按 CTRL+C 退出..."
        Read-Host | Out-Null
    }

    PrintHeader 'Custom Mode'
}


function ShowAppRemoval {
    PrintHeader "App Removal"

    $result = ShowAppSelectionForm

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Output "You have selected $($script:SelectedApps.Count) apps for removal"
        AddParameter 'RemoveAppsCustom' "Remove $($script:SelectedApps.Count) apps:"

        # 如果静音参数已通过，请抑制提示
        if (-not $Silent) {
            Write-Output ""
            Write-Output ""
            Write-Output "Press enter to remove the selected apps or press CTRL+C to quit..."
            Read-Host | Out-Null
            PrintHeader "App Removal"
        }
    }
    else {
        Write-Host "Selection was cancelled, no apps have been removed" -ForegroundColor Red
        Write-Output ""
    }
}


function LoadAndShowSavedSettings {
    PrintHeader 'Custom Mode'
    Write-Output "Win11Debloat will make the following changes:"

    # 从文件中打印已保存的设置信息
    Foreach ($line in (Get-Content -Path "$PSScriptRoot/SavedSettings" )) { 
        # 删除行前后的任何空格
        $line = $line.Trim()
    
        # 检查该行是否包含注释
        if (-not ($line.IndexOf('#') -eq -1)) {
            $parameterName = $line.Substring(0, $line.IndexOf('#'))

            # 打印参数描述并将参数添加到参数列表中
            switch ($parameterName) {
                'RemoveApps' {
                    PrintAppsList "$PSScriptRoot/Appslist.txt" $true
                }
                'RemoveAppsCustom' {
                    PrintAppsList "$PSScriptRoot/CustomAppsList" $true
                }
                default {
                    Write-Output $line.Substring(($line.IndexOf('#') + 1), ($line.Length - $line.IndexOf('#') - 1))
                }
            }

            if (-not $script:Params.ContainsKey($parameterName)) {
                $script:Params.Add($parameterName, $true)
            }
        }
    }

    # 如果静音参数已通过，请抑制提示
    if (-not $Silent) {
        Write-Output ""
        Write-Output ""
        Write-Output "按回车键执行脚本或按 CTRL+C 退出..."
        Read-Host | Out-Null
    }

    PrintHeader 'Custom Mode'
}



##################################################################################################################
#                                                                                                                #
#                                                  SCRIPT START                                                  #
#                                                                                                                #
##################################################################################################################



# 检查 winget 是否已安装,如果是.请检查版本是否至少为 v1.4
if ((Get-AppxPackage -Name "*Microsoft.DesktopAppInstaller*") -and ([int](((winget -v) -replace 'v','').split('.')[0..1] -join '') -gt 14)) {
    $script:wingetInstalled = $true
}
else {
    $script:wingetInstalled = $false

    # 顯示需要使用者確認的警告,如果沉默引數已透過,則抑制確認
    if (-not $Silent) {
        Write-Warning "Winget is not installed or outdated. This may prevent Win11Debloat from removing certain apps."
        Write-Output ""
        Write-Output "Press any key to continue anyway..."
        $null = [System.Console]::ReadKey()
    }
}

# 获取当前的 Windows 构建版本,以便与功能进行比较
$WinVersion = Get-ItemPropertyValue 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' CurrentBuild

# 检查机器是否支持现代待机,这用于确定是否可以使用 DisableModernStandbyNetworking 选项
$script:ModernStandbySupported = CheckModernStandbySupport

$script:Params = $PSBoundParameters
$script:FirstSelection = $true
$SPParams = 'WhatIf', 'Confirm', 'Verbose', 'Silent', 'Sysprep', 'Debug', 'User', 'CreateRestorePoint', 'LogPath'
$SPParamCount = 0

# 计算参数中存在多少个 SP 参数
# 这后来用于检查是否选择了任何选项
foreach ($Param in $SPParams) {
    if ($script:Params.ContainsKey($Param)) {
        $SPParamCount++
    }
}

# 隐藏应用程序删除的进度条,因为它们阻止了 Win11Debloat 的输出
if (-not ($script:Params.ContainsKey("Verbose"))) {
    $ProgressPreference = 'SilentlyContinue'
}
else {
    Write-Host "Verbose mode is enabled"
    Write-Output ""
    Write-Output "Press any key to continue..."
    $null = [System.Console]::ReadKey()

    $ProgressPreference = 'Continue'
}

if ($script:Params.ContainsKey("Sysprep")) {
    $defaultUserPath = GetUserDirectory -userName "Default"

    # 如果在 Sysprep 模式下运行,退出脚本 Windows 10
    if ($WinVersion -lt 22000) {
        Write-Host "Error: Win11Debloat Sysprep mode is not supported on Windows 10" -ForegroundColor Red
        AwaitKeyToExit
    }
}

# 如果指定了用户,请确保满足用户模式的所有要求
if ($script:Params.ContainsKey("User")) {
    $userPath = GetUserDirectory -userName $script:Params.Item("User")
}

# 如果已保存的设置文件存在且为空,请删除该文件
if ((Test-Path "$PSScriptRoot/SavedSettings") -and ([String]::IsNullOrWhiteSpace((Get-content "$PSScriptRoot/SavedSettings")))) {
    Remove-Item -Path "$PSScriptRoot/SavedSettings" -recurse
}

# 仅当 "RunAppsListGenerator" 参数传递给脚本时,才运行应用程序选择表单
if ($RunAppConfigurator -or $RunAppsListGenerator) {
    PrintHeader "自定义应用程序列表生成器"

    $result = ShowAppSelectionForm

    # 根据应用程序选择是否已保存或取消显示不同的消息
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Application selection window was closed without saving." -ForegroundColor Red
    }
    else {
        Write-Output "您的应用程序选择已保存到 "CustomAppsList" 文件中，该文件位于:"
        Write-Host "$PSScriptRoot" -ForegroundColor Yellow
    }

    AwaitKeyToExit
}

# 根据提供的参数或用户输入更改脚本执行
if ((-not $script:Params.Count) -or $RunDefaults -or $RunDefaultsLite -or $RunSavedSettings -or ($SPParamCount -eq $script:Params.Count)) {
    if ($RunDefaults -or $RunDefaultsLite) {
        $Mode = '1'
    }
    elseif ($RunSavedSettings) {
        if (-not (Test-Path "$PSScriptRoot/SavedSettings")) {
            PrintHeader 'Custom Mode'
            Write-Host "Error: No saved settings found, no changes were made" -ForegroundColor Red
            AwaitKeyToExit
        }

        $Mode = '4'
    }
    else {
        $Mode = ShowScriptMenuOptions 
    }

    # 根据模式添加执行参数
    switch ($Mode) {
        # 默认模式，加载默认值和应用程序删除选项
        '1' { 
            ShowDefaultMode
        }

        # 自定义模式，显示用户选择的所有可用选项
        '2' { 
            ShowCustomModeOptions
        }

        # 删除应用程序，根据用户选择删除应用程序
        '3' {
            ShowAppRemoval
        }

        # 从 "SavedSettings" 文件中加载上次使用的自定义选项
        '4' {
            LoadAndShowSavedSettings
        }
    }
}
else {
    PrintHeader '自定义模式'
}

# 如果 SPParams 中的键数等于 Params 中的键数,则未选择修改/更改
#  或由用户添加,脚本可以在不进行任何更改的情况下退出.
if (($SPParamCount -eq $script:Params.Keys.Count) -or (($script:Params.Keys.Count -eq 1) -and ($script:Params.Keys -contains 'CreateRestorePoint'))) {
    Write-Output "The script completed without making any changes."

    AwaitKeyToExit
}

# 如果传递了 CreateRestorePoint 参数,则创建系统恢复点
if ($script:Params.ContainsKey("创建恢复点")) {
    CreateSystemRestorePoint
}

# 执行所有选定/提供的参数
switch ($script:Params.Keys) {
    'RemoveApps' {
        $appsList = ReadAppslistFromFile "$PSScriptRoot/Appslist.txt" 
        Write-Output "> 删除默认选择 $($appsList.Count) apps..."
        RemoveApps $appsList
        continue
    }
    'RemoveAppsCustom' {
        if (-not (Test-Path "$PSScriptRoot/CustomAppsList")) {
            Write-Host "> 错误: 无法从文件中加载自定义应用程序列表，没有删除任何应用程序" -ForegroundColor Red
            Write-Output ""
            continue
        }
        
        $appsList = ReadAppslistFromFile "$PSScriptRoot/CustomAppsList"
        Write-Output "> 移除 $($appsList.Count) apps..."
        RemoveApps $appsList
        continue
    }
    'RemoveCommApps' {
        $appsList = 'Microsoft.windowscommunicationsapps', 'Microsoft.People'
        Write-Output "> Removing Mail, Calendar and People apps..."
        RemoveApps $appsList
        continue
    }
    'RemoveW11Outlook' {
        $appsList = 'Microsoft.OutlookForWindows'
        Write-Output "> 删除新的 Windows 版 Outlook 应用程序..."
        RemoveApps $appsList
        continue
    }
    'RemoveGamingApps' {
        $appsList = 'Microsoft.GamingApp', 'Microsoft.XboxGameOverlay', 'Microsoft.XboxGamingOverlay'
        Write-Output "> 移除与游戏相关的应用程序..."
        RemoveApps $appsList
        continue
    }
    'RemoveHPApps' {
        $appsList = 'AD2F1837.HPAIExperienceCenter', 'AD2F1837.HPJumpStarts', 'AD2F1837.HPPCHardwareDiagnosticsWindows', 'AD2F1837.HPPowerManager', 'AD2F1837.HPPrivacySettings', 'AD2F1837.HPSupportAssistant', 'AD2F1837.HPSureShieldAI', 'AD2F1837.HPSystemInformation', 'AD2F1837.HPQuickDrop', 'AD2F1837.HPWorkWell', 'AD2F1837.myHP', 'AD2F1837.HPDesktopSupportUtilities', 'AD2F1837.HPQuickTouch', 'AD2F1837.HPEasyClean', 'AD2F1837.HPConnectedMusic', 'AD2F1837.HPFileViewer', 'AD2F1837.HPRegistration', 'AD2F1837.HPWelcome', 'AD2F1837.HPConnectedPhotopoweredbySnapfish', 'AD2F1837.HPPrinterControl'
        Write-Output "> Removing HP apps..."
        RemoveApps $appsList
        continue
    }
    "ForceRemoveEdge" {
        ForceRemoveEdge
        continue
    }
    'DisableDVR' {
        RegImport "> Disabling Xbox game/screen recording..." "Disable_DVR.reg"
        continue
    }
    'DisableGameBarIntegration' {
        RegImport "> Disabling Game Bar integration..." "Disable_Game_Bar_Integration.reg"
        continue
    }
    'DisableTelemetry' {
        RegImport "> Disabling telemetry, diagnostic data, activity history, app-launch tracking and targeted ads..." "Disable_Telemetry.reg"
        continue
    }
    {$_ -in "DisableSuggestions", "DisableWindowsSuggestions"} {
        RegImport "> Disabling tips, tricks, suggestions and ads across Windows..." "Disable_Windows_Suggestions.reg"
        continue
    }
    'DisableEdgeAds' {
        RegImport "> Disabling ads, suggestions and the MSN news feed in Microsoft Edge..." "Disable_Edge_Ads_And_Suggestions.reg"
        continue
    }
    {$_ -in "DisableLockscrTips", "DisableLockscreenTips"} {
        RegImport "> Disabling tips & tricks on the lockscreen..." "Disable_Lockscreen_Tips.reg"
        continue
    }
    'DisableDesktopSpotlight' {
        RegImport "> Disabling the 'Windows Spotlight' desktop background option..." "Disable_Desktop_Spotlight.reg"
        continue
    }
    'DisableSettings365Ads' {
        RegImport "> Disabling Microsoft 365 ads in Settings Home..." "Disable_Settings_365_Ads.reg"
        continue
    }
    'DisableSettingsHome' {
        RegImport "> Disabling the Settings Home page..." "Disable_Settings_Home.reg"
        continue
    }
    {$_ -in "DisableBingSearches", "DisableBing"} {
        RegImport "> Disabling Bing web search, Bing AI and Cortana from Windows search..." "Disable_Bing_Cortana_In_Search.reg"
        
        # Also remove the app package for Bing search
        $appsList = 'Microsoft.BingSearch'
        RemoveApps $appsList
        continue
    }
    'DisableCopilot' {
        RegImport "> Disabling Microsoft Copilot..." "Disable_Copilot.reg"

        # Also remove the app package for Copilot
        $appsList = 'Microsoft.Copilot'
        RemoveApps $appsList
        continue
    }
    'DisableRecall' {
        RegImport "> Disabling Windows Recall..." "Disable_AI_Recall.reg"
        continue
    }
    'DisableClickToDo' {
        RegImport "> Disabling Click to Do..." "Disable_Click_to_Do.reg"
        continue
    }
    'DisableEdgeAI' {
        RegImport "> Disabling AI features in Microsoft Edge..." "Disable_Edge_AI_Features.reg"
        continue
    }
    'DisablePaintAI' {
        RegImport "> Disabling AI features in Paint..." "Disable_Paint_AI_Features.reg"
        continue
    }
    'DisableNotepadAI' {
        RegImport "> Disabling AI features in Notepad..." "Disable_Notepad_AI_Features.reg"
        continue
    }
    'RevertContextMenu' {
        RegImport "> Restoring the old Windows 10 style context menu..." "Disable_Show_More_Options_Context_Menu.reg"
        continue
    }
    'DisableMouseAcceleration' {
        RegImport "> Turning off Enhanced Pointer Precision..." "Disable_Enhance_Pointer_Precision.reg"
        continue
    }
    'DisableStickyKeys' {
        RegImport "> Disabling the Sticky Keys keyboard shortcut..." "Disable_Sticky_Keys_Shortcut.reg"
        continue
    }
    'DisableFastStartup' {
        RegImport "> Disabling Fast Start-up..." "Disable_Fast_Startup.reg"
        continue
    }
    'DisableModernStandbyNetworking' {
        RegImport "> Disabling network connectivity during Modern Standby..." "Disable_Modern_Standby_Networking.reg"
        continue
    }
    'ClearStart' {
        Write-Output "> Removing all pinned apps from the start menu for user $(GetUserName)..."
        ReplaceStartMenu
        Write-Output ""
        continue
    }
    'ReplaceStart' {
        Write-Output "> Replacing the start menu for user $(GetUserName)..."
        ReplaceStartMenu $script:Params.Item("ReplaceStart")
        Write-Output ""
        continue
    }
    'ClearStartAllUsers' {
        ReplaceStartMenuForAllUsers
        continue
    }
    'ReplaceStartAllUsers' {
        ReplaceStartMenuForAllUsers $script:Params.Item("ReplaceStartAllUsers")
        continue
    }
    'DisableStartRecommended' {
        RegImport "> Disabling the start menu recommended section..." "Disable_Start_Recommended.reg"
        continue
    }
    'DisableStartPhoneLink' {
        RegImport "> Disabling the Phone Link mobile devices integration in the start menu..." "Disable_Phone_Link_In_Start.reg"
        continue
    }
    'EnableDarkMode' {
        RegImport "> Enabling dark mode for system and apps..." "Enable_Dark_Mode.reg"
        continue
    }
    'DisableTransparency' {
        RegImport "> Disabling transparency effects..." "Disable_Transparency.reg"
        continue
    }
    'DisableAnimations' {
        RegImport "> Disabling animations and visual effects..." "Disable_Animations.reg"
        continue
    }
    'TaskbarAlignLeft' {
        RegImport "> Aligning taskbar buttons to the left..." "Align_Taskbar_Left.reg"
        continue
    }
    'CombineTaskbarAlways' {
        RegImport "> Setting the taskbar on the main display to always combine buttons and hide labels..." "Combine_Taskbar_Always.reg"
        continue
    }
    'CombineTaskbarWhenFull' {
        RegImport "> Setting the taskbar on the main display to only combine buttons and hide labels when the taskbar is full..." "Combine_Taskbar_When_Full.reg"
        continue
    }
    'CombineTaskbarNever' {
        RegImport "> Setting the taskbar on the main display to never combine buttons or hide labels..." "Combine_Taskbar_Never.reg"
        continue
    }
    'CombineMMTaskbarAlways' {
        RegImport "> Setting the taskbar on secondary displays to always combine buttons and hide labels..." "Combine_MMTaskbar_Always.reg"
        continue
    }
    'CombineMMTaskbarWhenFull' {
        RegImport "> Setting the taskbar on secondary displays to only combine buttons and hide labels when the taskbar is full..." "Combine_MMTaskbar_When_Full.reg"
        continue
    }
    'CombineMMTaskbarNever' {
        RegImport "> Setting the taskbar on secondary displays to never combine buttons or hide labels..." "Combine_MMTaskbar_Never.reg"
        continue
    }
    'MMTaskbarModeAll' {
        RegImport "> Setting the taskbar to only show app icons on main taskbar..." "MMTaskbarMode_All.reg"
        continue
    }
    'MMTaskbarModeMainActive' {
        RegImport "> Setting the taskbar to show app icons on all taskbars..." "MMTaskbarMode_Main_Active.reg"
        continue
    }
    'MMTaskbarModeActive' {
        RegImport "> Setting the taskbar to only show app icons on the taskbar where the window is open..." "MMTaskbarMode_Active.reg"
        continue
    }
    'HideSearchTb' {
        RegImport "> Hiding the search icon from the taskbar..." "Hide_Search_Taskbar.reg"
        continue
    }
    'ShowSearchIconTb' {
        RegImport "> Changing taskbar search to icon only..." "Show_Search_Icon.reg"
        continue
    }
    'ShowSearchLabelTb' {
        RegImport "> Changing taskbar search to icon with label..." "Show_Search_Icon_And_Label.reg"
        continue
    }
    'ShowSearchBoxTb' {
        RegImport "> Changing taskbar search to search box..." "Show_Search_Box.reg"
        continue
    }
    'HideTaskview' {
        RegImport "> Hiding the taskview button from the taskbar..." "Hide_Taskview_Taskbar.reg"
        continue
    }
    {$_ -in "HideWidgets", "DisableWidgets"} {
        RegImport "> Disabling widgets on the taskbar & lockscreen..." "Disable_Widgets_Service.reg"

        # Also remove the app package for Widgets
        $appsList = 'Microsoft.StartExperiencesApp'
        RemoveApps $appsList
        continue
    }
    {$_ -in "HideChat", "DisableChat"} {
        RegImport "> Hiding the chat icon from the taskbar..." "Disable_Chat_Taskbar.reg"
        continue
    }
    'EnableEndTask' {
        RegImport "> Enabling the 'End Task' option in the taskbar right click menu..." "Enable_End_Task.reg"
        continue
    }
    'EnableLastActiveClick' {
        RegImport "> Enabling the 'Last Active Click' behavior in the taskbar app area..." "Enable_Last_Active_Click.reg"
        continue
    }
    'ExplorerToHome' {
        RegImport "> Changing the default location that File Explorer opens to `Home`..." "Launch_File_Explorer_To_Home.reg"
        continue
    }
    'ExplorerToThisPC' {
        RegImport "> Changing the default location that File Explorer opens to `This PC`..." "Launch_File_Explorer_To_This_PC.reg"
        continue
    }
    'ExplorerToDownloads' {
        RegImport "> Changing the default location that File Explorer opens to `Downloads`..." "Launch_File_Explorer_To_Downloads.reg"
        continue
    }
    'ExplorerToOneDrive' {
        RegImport "> Changing the default location that File Explorer opens to `OneDrive`..." "Launch_File_Explorer_To_OneDrive.reg"
        continue
    }
    'ShowHiddenFolders' {
        RegImport "> Unhiding hidden files, folders and drives..." "Show_Hidden_Folders.reg"
        continue
    }
    'ShowKnownFileExt' {
        RegImport "> Enabling file extensions for known file types..." "Show_Extensions_For_Known_File_Types.reg"
        continue
    }
    'HideHome' {
        RegImport "> Hiding the home section from the File Explorer navigation pane..." "Hide_Home_from_Explorer.reg"
        continue
    }
    'HideGallery' {
        RegImport "> Hiding the gallery section from the File Explorer navigation pane..." "Hide_Gallery_from_Explorer.reg"
        continue
    }
    'HideDupliDrive' {
        RegImport "> Hiding duplicate removable drive entries from the File Explorer navigation pane..." "Hide_duplicate_removable_drives_from_navigation_pane_of_File_Explorer.reg"
        continue
    }
    {$_ -in "HideOnedrive", "DisableOnedrive"} {
        RegImport "> Hiding the OneDrive folder from the File Explorer navigation pane..." "Hide_Onedrive_Folder.reg"
        continue
    }
    {$_ -in "Hide3dObjects", "Disable3dObjects"} {
        RegImport "> Hiding the 3D objects folder from the File Explorer navigation pane..." "Hide_3D_Objects_Folder.reg"
        continue
    }
    {$_ -in "HideMusic", "DisableMusic"} {
        RegImport "> Hiding the music folder from the File Explorer navigation pane..." "Hide_Music_folder.reg"
        continue
    }
    {$_ -in "HideIncludeInLibrary", "DisableIncludeInLibrary"} {
        RegImport "> Hiding 'Include in library' in the context menu..." "Disable_Include_in_library_from_context_menu.reg"
        continue
    }
    {$_ -in "HideGiveAccessTo", "DisableGiveAccessTo"} {
        RegImport "> Hiding 'Give access to' in the context menu..." "Disable_Give_access_to_context_menu.reg"
        continue
    }
    {$_ -in "HideShare", "DisableShare"} {
        RegImport "> Hiding 'Share' in the context menu..." "Disable_Share_from_context_menu.reg"
        continue
    }
}

RestartExplorer

Write-Output ""
Write-Output ""
Write-Output ""
Write-Output "脚本完成! 请检查上面是否有任何错误."

AwaitKeyToExit
