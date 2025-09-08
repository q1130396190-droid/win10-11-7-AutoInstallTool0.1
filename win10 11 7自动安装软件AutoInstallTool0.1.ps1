﻿# 检查管理员权限
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Definition)`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName System.Windows.Forms

# 创建主表单
$form = New-Object System.Windows.Forms.Form
$form.Text = "软件安装与系统优化"
$form.Width = 600
$form.Height = 400

# 系统信息显示框
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Width = 550
$listBox.Height = 200
$listBox.Location = New-Object System.Drawing.Point(25, 20)
$form.Controls.Add($listBox)

# 进度条
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Width = 550
$progressBar.Height = 20
$progressBar.Location = New-Object System.Drawing.Point(25, 250)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$form.Controls.Add($progressBar)

# 状态标签
$label = New-Object System.Windows.Forms.Label
$label.Text = "正在初始化..."
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(25, 280)
$form.Controls.Add($label)

# 显示窗口
$form.Add_Shown({$form.Activate()})
$form.Show()

# 系统检测函数
function Get-OSVersion {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $version = [System.Environment]::OSVersion.Version
    
    if ($version.Major -eq 6 -and $version.Minor -eq 1) {
        return "Windows 7"
    }
    elseif ($version.Major -eq 10 -and $version.Build -ge 22000) {
        return "Windows 11"
    }
    return "Unknown"
}

# 系统优化函数
function Optimize-System {
    param($osType)
    
    $listBox.Items.Add("开始系统优化...")
    $label.Text = "正在应用 $osType 优化设置"
    $form.Refresh()

    switch ($osType) {
        "Windows 7" {
            # 禁用UAC
            Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -Value 0 -ErrorAction SilentlyContinue
            
            # 最佳性能设置
            Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "DragFullWindows" -Value "0"
            Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "FontSmoothing" -Value "0"
            powercfg -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c  # 高性能电源计划
        }
        "Windows 11" {
            # 禁用动画效果
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarAnimations" -Value 0
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ListviewAlphaSelect" -Value 0
            
            # 优化隐私设置
            Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0
        }
        default {
            $listBox.Items.Add("未知系统类型，跳过优化步骤")
        }
    }
    $listBox.Items.Add("系统优化完成")
}

# 检测系统版本
$osType = Get-OSVersion
$listBox.Items.Add("检测到操作系统: $osType")
$form.Refresh()

# 执行系统优化
Optimize-System -osType $osType

# 安装软件部分（添加@符号过滤）
$currentDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
# 获取所有EXE文件，但排除文件名包含@的文件
$files = Get-ChildItem -Path $currentDirectory -Recurse -Filter "*.exe" | Where-Object { 
    $_.Name -notlike "*@*"  # 核心过滤条件：文件名不包含@
}
$totalFiles = $files.Count
$currentFileIndex = 0

# 显示检测到的安装包
$listBox.Items.Add("发现 $totalFiles 个安装程序（已排除含@的文件）:")
foreach ($file in $files) {
    $listBox.Items.Add("-> $($file.Name)")
}

# 安装循环
foreach ($file in $files) {
    $currentFileIndex++
    $progress = [math]::Round(($currentFileIndex / $totalFiles) * 100)
    $label.Text = "正在安装: $($file.Name) ($progress%)"
    $progressBar.Value = $progress
    $form.Refresh()

    # 根据系统类型选择安装参数
    $arguments = switch ($osType) {
        "Windows 7" { "/S", "/norestart" }
        "Windows 11" { "/silent", "/norestart" }
        default { "/quiet", "/norestart" }
    }

    try {
        $process = Start-Process -FilePath $file.FullName -ArgumentList $arguments -PassThru -NoNewWindow -ErrorAction Stop
        $process | Wait-Process -Timeout 300 -ErrorAction Stop  # 5分钟超时
        $listBox.Items.Add("安装成功: $($file.Name)")
    }
    catch {
        $listBox.Items.Add("安装失败: $($file.Name) - $($_.Exception.Message)")
    }
    finally {
        if ($process -and -not $process.HasExited) {
            $process.Kill()
        }
    }
}

# 完成处理
$label.Text = "所有操作已完成"
$progressBar.Value = 100
[System.Windows.Forms.MessageBox]::Show("所有操作已完成", "完成", "OK", "Information")
$form.Close()