# =================================================================================
# PowerShell 脚本：从指定 URL 下载文件
#
# 功能:
#   - 自动将文件下载到 PowerShell 当前所在的目录。
#   - 如果文件已存在，则跳过下载。
#   - 为重定向的下载链接指定一个清晰的文件名。
#   - 提供下载进度和错误处理。
#
# 要求:
#   - 建议以管理员身份运行，以避免权限问题。
# =================================================================================

# --- 1. 准备工作 ---

<#
# 强制使用 TLS 1.2 协议，提高与 GitHub 等现代网站的兼容性
# 这是一个重要的最佳实践，可以避免很多潜在的连接错误
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}
catch {
    Write-Warning "无法设置 TLS 1.2。在旧版 PowerShell 上可能会遇到连接问题。"
}
#>


# 自动获取脚本文件所在的目录路径
# $PSScriptRoot 是一个自动变量，代表当前脚本所在的文件夹
# --- 关键修改：替换 $PSScriptRoot ---
# 在通过 iex 进行网络执行时，$PSScriptRoot 是空的，会导致错误。
# 我们使用 $PWD.Path 来替代它，$PWD 代表 PowerShell 当前的工作目录 (Current Working Directory)。
# 这样，文件就会被下载到你执行命令时所在的那个文件夹里。
$destinationFolder = $PWD.Path

Write-Host "文件将被下载到目录: $destinationFolder" -ForegroundColor Cyan


# --- 2. 定义要下载的文件列表 ---

# 我们使用一个哈希表来存储 URL 和对应的输出文件名
# 这样可以优雅地处理没有直接文件名的下载链接
$filesToDownload = @{
    "https://github.com/work-reporter/work-reports/raw/refs/heads/main/Soft/WinRAR%20v7.13%20x64%20SC.exe" = "WinRAR v7.13 x64 SC.exe";
}


# --- 3. 循环遍历并执行下载 ---

# .GetEnumerator() 让我们可以在哈希表上进行循环
foreach ($fileEntry in $filesToDownload.GetEnumerator()) {
    
    $url = $fileEntry.Key
    $fileName = $fileEntry.Value
    
    # 组合成完整的目标文件路径
    $destinationPath = Join-Path -Path $destinationFolder -ChildPath $fileName
    
    Write-Host ("-" * 50)
    
    # 检查文件是否已存在
    if (Test-Path $destinationPath) {
        Write-Host "文件已存在，跳过下载: $fileName" -ForegroundColor Yellow
        continue # 跳到循环的下一次迭代
    }
    
    # 使用 try...catch 结构来捕获可能发生的下载错误
    try {
        Write-Host "正在下载: $fileName" -ForegroundColor Green
        Write-Host "来源 URL: $url"
        
        # 执行下载命令，-Uri 是源地址，-OutFile 是保存路径
        # 在 PowerShell 控制台中运行时，此命令会自动显示一个进度条
        Invoke-WebRequest -Uri $url -OutFile $destinationPath
        
        Write-Host "成功保存到: $destinationPath" -ForegroundColor Green
    }
    catch {
        # 如果下载过程中发生任何错误，则执行此处的代码
        Write-Error "下载文件 '$fileName' 时发生错误!"
        Write-Error "错误详情: $_" # $_ 包含了详细的错误信息
    }
}

Write-Host ("-" * 50)
Write-Host "所有下载任务已处理完毕。" -ForegroundColor Cyan


# --- 4. 软件安装 (Software Installation) ---

# 定义软件安装包所在的目录 (这里也需要使用 $destinationFolder)
$softwarePath = $destinationFolder
# 假设WinRAR的安装文件名为 "winrar-x64.exe"，请根据实际情况修改
$winrarInstaller = Join-Path $softwarePath "WinRAR v7.13 x64 SC.exe"

Write-Host "--- 开始安装软件 ---"

# 检查WinRAR安装文件是否存在
if (Test-Path $winrarInstaller) {
    Write-Host "正在静默安装 WinRAR..."
    # 使用 /S 参数进行完全静默安装
    Start-Process $winrarInstaller -ArgumentList "/S" -Wait
    Write-Host "WinRAR 安装完成。"
    # 安装完成后删除安装包
    Remove-Item -Path $winrarInstaller -Force
} else {
    Write-Host "警告: 未在 $winrarInstaller 找到WinRAR安装文件，跳过安装。"
}

Write-Host "--- 软件安装阶段结束 ---`n"
# 在自动化脚本中，最后的 Read-Host 通常可以移除，除非你确实需要暂停
# Read-Host -Prompt "按 Enter 键退出脚本..."