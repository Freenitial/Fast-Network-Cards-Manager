<# :
    REM Author: Leo Gillet - Freenitial on GitHub
    @echo off & chcp 437 >nul & Title Fast Network Cards Manager
    copy /y "%~f0" "%TEMP%\%~n0.ps1" >NUL && powershell -Nologo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%TEMP%\%~n0.ps1"
    exit /b
#>

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Restart with admin rights
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`"" -Verb RunAs 
    Exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class DPIHelper {
    public static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);
    [DllImport("user32.dll")]
    public static extern bool SetProcessDpiAwarenessContext(IntPtr dpiFlag);
}
"@ -ReferencedAssemblies @("System.Runtime.InteropServices")
[DPIHelper]::SetProcessDpiAwarenessContext([DPIHelper]::DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2) | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "Fast Network Cards Manager"
$form.Size = New-Object System.Drawing.Size(700,350)
$form.StartPosition = "CenterScreen"
$form.AutoScaleMode = "Dpi"
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayout.Dock = "Fill"
$tableLayout.ColumnCount = 1
$tableLayout.RowCount = 2
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,75))) | Out-Null
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,100))) | Out-Null
$form.Controls.Add($tableLayout)

$listView = New-Object System.Windows.Forms.ListView
$listView.Sorting = [System.Windows.Forms.SortOrder]::Ascending
$listView.View = 'Details'
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.CheckBoxes = $true
$listView.Dock = "Fill"

$colDesc = New-Object System.Windows.Forms.ColumnHeader
$colDesc.Text = "Description"
$colDesc.Width = 500
$colStatus = New-Object System.Windows.Forms.ColumnHeader
$colStatus.Text = "Status"
$colStatus.Width = 150

$listView.Columns.AddRange(@($colDesc, $colStatus))
$tableLayout.Controls.Add($listView, 0, 0)

$bottomPanel = New-Object System.Windows.Forms.TableLayoutPanel
$bottomPanel.Dock = "Fill"
$bottomPanel.ColumnCount = 5
$bottomPanel.RowCount = 2
$bottomPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$bottomPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,20))) | Out-Null
for ($i=0; $i -lt 5; $i++) {
    $bottomPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,20))) | Out-Null
}
$tableLayout.Controls.Add($bottomPanel, 0, 1)

$buttonEnableSelection = New-Object System.Windows.Forms.Button
$buttonEnableSelection.Text = "Enable Selection"
$buttonEnableSelection.Dock = "Fill"
$bottomPanel.Controls.Add($buttonEnableSelection, 0, 0)

$buttonDisableSelection = New-Object System.Windows.Forms.Button
$buttonDisableSelection.Text = "Disable Selection"
$buttonDisableSelection.Dock = "Fill"
$bottomPanel.Controls.Add($buttonDisableSelection, 1, 0)

$buttonEnableAll = New-Object System.Windows.Forms.Button
$buttonEnableAll.Text = "Enable All"
$buttonEnableAll.Dock = "Fill"
$bottomPanel.Controls.Add($buttonEnableAll, 2, 0)

$buttonDisableAll = New-Object System.Windows.Forms.Button
$buttonDisableAll.Text = "Disable All"
$buttonDisableAll.Dock = "Fill"
$bottomPanel.Controls.Add($buttonDisableAll, 3, 0)

$buttonRefresh = New-Object System.Windows.Forms.Button
$buttonRefresh.Text = "Refresh Adapters"
$buttonRefresh.Dock = "Fill"
$bottomPanel.Controls.Add($buttonRefresh, 4, 0)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Dock = "Fill"
$bottomPanel.SetColumnSpan($progressBar, 5)
$bottomPanel.Controls.Add($progressBar, 0, 0)

$script:adapterQueue = @()    # Queue of adapter objects to process
$script:currentOperation = "" # "Disable" or "Enable"
$script:progressValue = 0     # Current progress value
$script:progressStep = 0      # Increment per adapter

function Reset-AdapterList {
    $listView.Items.Clear()
    try {
        $adapters = Get-NetAdapter -IncludeHidden
        foreach ($adapter in $adapters) {
            if ([string]::IsNullOrEmpty($adapter.InterfaceDescription) -or $adapter.InterfaceDescription -eq "Microsoft Kernel Debug Network Adapter") { continue }
            $item = New-Object System.Windows.Forms.ListViewItem($adapter.InterfaceDescription)
            $item.SubItems.Add($adapter.Status) | Out-Null
            $item.ForeColor = if ($adapter.Status -eq "Up") { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Red }
            $item.Tag = $adapter
            $listView.Items.Add($item) | Out-Null
        }
        $listView.Sort()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error retrieving network adapters: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}
$buttonRefresh.Add_Click({ Reset-AdapterList })

function Update-ProgressBar {
    param([int]$progress)
    if ($progress -gt 100) { $progress = 100 }
    $form.Invoke([action]{ $progressBar.Value = $progress }) | Out-Null
}

# Timer to process each adapter and update progressbar (single thread)
$processTimer = New-Object System.Windows.Forms.Timer
$processTimer.Interval = 100
$processTimer.add_Tick({
    if ($script:adapterQueue.Count -gt 0) {
        $adapter = $script:adapterQueue[0]
        if ($script:adapterQueue.Count -gt 1) { $script:adapterQueue = $script:adapterQueue[1..($script:adapterQueue.Count - 1)] }
        else { $script:adapterQueue = @() }
        $newProgress = $script:progressValue + $script:progressStep
        if ($newProgress -gt 100) { $newProgress = 100 }
        try {
            if ($script:currentOperation -eq "Disable") { Disable-NetAdapter -IncludeHidden -Name $adapter.Name -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
            elseif ($script:currentOperation -eq "Enable") { Enable-NetAdapter -IncludeHidden -Name $adapter.Name -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
        }
        catch { } # Ignore errors
        $script:progressValue = $newProgress
        Update-ProgressBar $script:progressValue
    }
    else {
        $processTimer.Stop()
        Update-ProgressBar 100
        Reset-AdapterList
    }
})

function Update-Adapters {
    param(
        [ScriptBlock]$Filter,
        [string]$Operation, # "Enable" or "Disable"
        [string]$ErrorMessage
    )
    $targetAdapters = @()
    foreach ($item in $listView.Items) {
        $adapter = $item.Tag
        if (& $Filter $adapter) { $targetAdapters += $adapter }
    }
    if ($targetAdapters.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show($ErrorMessage, "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $script:adapterQueue = $targetAdapters
    $script:currentOperation = $Operation
    $script:progressValue = 0
    $script:progressStep = [math]::Floor(100 / $targetAdapters.Count)
    $progressBar.Value = 0
    $processTimer.Start()
}

$buttonDisableAll.Add_Click({ Update-Adapters { param($adapter) $adapter.Status -eq "Up" } "Disable" "No active network adapters found." })
$buttonEnableAll.Add_Click({ Update-Adapters { param($adapter) $adapter.Status -ne "Up" } "Enable" "No disabled network adapters found." })
$buttonDisableSelection.Add_Click({
    Update-Adapters {
        param($adapter)
        $item = $listView.Items | Where-Object { $_.Tag -eq $adapter }
        $item.Checked -and $adapter.Status -eq "Up"
    } "Disable" "No applicable selected adapters found."
})
$buttonEnableSelection.Add_Click({
    Update-Adapters {
        param($adapter)
        $item = $listView.Items | Where-Object { $_.Tag -eq $adapter }
        $item.Checked -and $adapter.Status -ne "Up"
    } "Enable" "No applicable selected adapters found."
})

Reset-AdapterList
[System.Windows.Forms.Application]::Run($form)
if ($PSCommandPath) { Remove-Item $PSCommandPath -Force }
