param (
    [switch]$Silent,
    [switch]$Verbose,
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

# 如果当前的powershell环境没有将语言模式设置为全语言，则显示错误
if ($ExecutionContext.SessionState.LanguageMode -ne "FullLanguage") {
   Write-Host "错误：Win11Debloat无法在您的系统上运行。 PowerShell的执行受到安全策略的限制" -ForegroundColor Red
   Write-Output ""
   Write-Output "按回车键退出..."
   Read-Host | Out-Null
   Exit
}

Clear-Host
Write-Output "-------------------------------------------------------------------------------------------"
Write-Output " Win11Debloat脚本-获取"
Write-Output "-------------------------------------------------------------------------------------------"

Write-Output "> 正在下载Win11Debloat..."

# 从GitHub下载最新版本的Win11Debloat作为zip存档
try {
    $LatestReleaseUri = (Invoke-RestMethod https://api.github.com/repos/Raphire/Win11Debloat/releases/latest).zipball_url
    Invoke-RestMethod $LatestReleaseUri -OutFile "$env:TEMP/win11debloat.zip"
}
catch {
    Write-Host "错误：无法从GitHub获取最新版本。 请检查您的互联网连接，然后重试." -ForegroundColor Red
    Write-Output ""
    Write-Output "按回车键退出..."
    Read-Host | Out-Null
    Exit
}

# 删除旧脚本文件夹（如果存在），CustomAppsList和SavedSettings文件除外
if (Test-Path "$env:TEMP/Win11Debloat") {
    Write-Output ""
    Write-Output "> 清理旧的Win11Debloat文件夹..."
    Get-ChildItem -Path "$env:TEMP/Win11Debloat" -Exclude CustomAppsList,SavedSettings,Win11Debloat.log | Remove-Item -Recurse -Force
}

Write-Output ""
Write-Output "> 正在解压..."

# 将存档解压缩到Win11Debloat文件夹
Expand-Archive "$env:TEMP/win11debloat.zip" "$env:TEMP/Win11Debloat"

# 删除存档
Remove-Item "$env:TEMP/win11debloat.zip"

# 移动文件
Get-ChildItem -Path "$env:TEMP/Win11Debloat/Raphire-Win11Debloat-*" -Recurse | Move-Item -Destination "$env:TEMP/Win11Debloat"

# 列出要传递到脚本的参数列表
$arguments = $($PSBoundParameters.GetEnumerator() | ForEach-Object {
    if ($_.Value -eq $true) {
        "-$($_.Key)"
    } 
    else {
         "-$($_.Key) ""$($_.Value)"""
    }
})

Write-Output ""
Write-Output "> 运行Win11Debloat..."

# 使用提供的参数运行Win11Debloat脚本
$debloatProcess = Start-Process powershell.exe -PassThru -ArgumentList "-executionpolicy bypass -File $env:TEMP\Win11Debloat\Win11Debloat.ps1 $arguments" -Verb RunAs

# 等待流程完成后再继续
if ($null -ne $debloatProcess) {
    $debloatProcess.WaitForExit()
}

# 删除所有剩余的脚本文件，CustomAppsList和SavedSettings文件除外
if (Test-Path "$env:TEMP/Win11Debloat") {
    Write-Output ""
    Write-Output "> Cleaning up..."

    # 清理，删除Win11Debloat目录
    Get-ChildItem -Path "$env:TEMP/Win11Debloat" -Exclude CustomAppsList,SavedSettings,Win11Debloat.log | Remove-Item -Recurse -Force
}

Write-Output ""
