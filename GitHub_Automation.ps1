# Filename: GitHub_Automation.ps1
# Purpose: Script tự động upload code lên GitHub với tính năng kiểm tra thay đổi (chỉ upload file có sự khác biệt).
# Token: (Removed for security, use GITHUB_TOKEN environment variable)

# Đảm bảo sử dụng TLS 1.2 cho kết nối an toàn với GitHub API
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$token = $env:GITHUB_TOKEN
$owner = "buiquangtrung2012-ops"
$repo = "HoSoAddin"
$branch = "main"

# Hàm tính Git SHA của file local (giống cách Git tính)
function Get-GitBlobSha {
    param($filePath)
    $content = [System.IO.File]::ReadAllBytes($filePath)
    $header = "blob $($content.Length)`0"
    $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
    $combinedBytes = New-Object byte[] ($headerBytes.Length + $content.Length)
    [Buffer]::BlockCopy($headerBytes, 0, $combinedBytes, 0, $headerBytes.Length)
    [Buffer]::BlockCopy($content, 0, $combinedBytes, $headerBytes.Length, $content.Length)
    
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    $hash = $sha1.ComputeHash($combinedBytes)
    return [System.BitConverter]::ToString($hash).Replace("-", "").ToLower()
}

function Update-GitHubFileApi {
    param ($filePath, $message, $current, $total, $remoteSha)
    
    $fullPath = (Resolve-Path $filePath).Path
    $basePath = (Get-Location).Path
    
    if ($fullPath.StartsWith($basePath)) {
        $relativePath = $fullPath.Substring($basePath.Length).TrimStart("\").Replace("\", "/")
    } else {
        $relativePath = Split-Path $fullPath -Leaf
    }
    
    # Tính SHA local
    $localSha = Get-GitBlobSha -filePath $fullPath
    
    # Nếu SHA khớp, bỏ qua upload
    if ($localSha -eq $remoteSha) {
        Write-Host "[$current/$total] Bo qua (khong doi): $relativePath" -ForegroundColor Gray
        return
    }

    $url = "https://api.github.com/repos/$owner/$repo/contents/$relativePath"
    $headers = @{
        "Authorization" = "token $token"
        "Accept" = "application/vnd.github.v3+json"
    }
    
    # Đọc dữ liệu binary và chuyển sang Base64
    $contentBytes = [System.IO.File]::ReadAllBytes($fullPath)
    $contentBase64 = [Convert]::ToBase64String($contentBytes)
    
    $body = @{
        message = $message
        content = $contentBase64
        branch = $branch
    }
    if ($remoteSha) { $body.sha = $remoteSha }
    
    $bodyJson = $body | ConvertTo-Json -Compress
    
    Write-Host "[$current/$total] Cloud Upload: $relativePath ..." -NoNewline -ForegroundColor Cyan
    try {
        $apiRes = Invoke-RestMethod -Uri $url -Headers $headers -Method Put -Body $bodyJson
        Write-Host " [OK]" -ForegroundColor Green
    } catch {
        Write-Host " [LOI]" -ForegroundColor Red
        if ($_.Exception.Response) {
            $stream = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            $errorBody = $reader.ReadToEnd()
            Write-Host "    Chi tiet: $errorBody" -ForegroundColor Gray
        } else {
            Write-Host "    Loi: $($_.Exception.Message)" -ForegroundColor Gray
        }
    }
}

function Push-ToGitHub {
    param (
        [string]$Message = "Automated Push"
    )

    $v_tag = "Update"
    if (Test-Path "taskpane.html") {
        $html = Get-Content "taskpane.html" -Raw -Encoding UTF8
        if ($html -match 'v(\d{8}\.\d{4})') {
            $v_tag = $matches[0]
        }
    }
    $finalMessage = "$v_tag - $Message"

    Write-Host ">>> Dang chuan bi upload version: $v_tag" -ForegroundColor White

    # Kiểm tra xem Git CLI có sẵn không
    $gitExists = Get-Command git -ErrorAction SilentlyContinue
    
    if ($gitExists) {
        Write-Host ">>> Su dung Git CLI (Tự động nhận diện thay đổi)..." -ForegroundColor Yellow
        git add .
        git commit -m $finalMessage
        git push origin main
        if ($LASTEXITCODE -ne 0) {
            git push origin master
        }
    } else {
        Write-Host ">>> Git CLI khong tim thay. Dang dung GitHub API Fallback..." -ForegroundColor Magenta
        
        # Lấy bản đồ SHA của toàn bộ repo trong 1 lần gọi để tối ưu tốc độ
        Write-Host ">>> Dang kiem tra trang thai file tren GitHub..." -ForegroundColor Gray
        $headers = @{ "Authorization" = "token $token"; "Accept" = "application/vnd.github.v3+json" }
        $treeUrl = "https://api.github.com/repos/$owner/$repo/git/trees/$branch?recursive=1"
        $remoteFiles = @{}
        try {
            $treeRes = Invoke-RestMethod -Uri $treeUrl -Headers $headers -Method Get
            if ($treeRes.tree) {
                foreach ($item in $treeRes.tree) {
                    if ($item.type -eq "blob") {
                        $remoteFiles[$item.path] = $item.sha
                    }
                }
            }
        } catch {
            Write-Host "Warning: Khong the lay tree tu GitHub. Se dung fallback tung file." -ForegroundColor Yellow
        }

        # Lọc các file cần upload
        $files = Get-ChildItem -File -Recurse | Where-Object { 
            $_.FullName -notmatch "\\\.git\\" -and 
            $_.FullName -notmatch "\\node_modules\\" -and
            $_.Extension -match "\.html$|\.js$|\.xml$|\.md$|\.bat$|\.ps1$|\.css$"
        }
        
        $totalFiles = $files.Count
        $i = 0
        foreach ($file in $files) {
            $i++
            $fullPath = $file.FullName
            $basePath = (Get-Location).Path
            $relPath = $fullPath.Substring($basePath.Length).TrimStart("\").Replace("\", "/")
            
            $rSha = $remoteFiles[$relPath]
            
            # Fallback nếu Tree API thất bại: Thử lấy SHA của từng file
            if (-not $rSha) {
                try {
                    $fileUrl = "https://api.github.com/repos/$owner/$repo/contents/$relPath"
                    $fileRes = Invoke-RestMethod -Uri $fileUrl -Headers $headers -Method Get -ErrorAction SilentlyContinue
                    $rSha = $fileRes.sha
                } catch { }
            }
            
            Update-GitHubFileApi -filePath $fullPath -message $finalMessage -current $i -total $totalFiles -remoteSha $rSha
        }
    }
    
    Write-Host ""
    Write-Host "============================" -ForegroundColor Green
    Write-Host "   UPLOAD HOAN TAT!         " -ForegroundColor Green
    Write-Host "============================" -ForegroundColor Green
}

# RUN
Push-ToGitHub -Message "Cập nhật từ local"
