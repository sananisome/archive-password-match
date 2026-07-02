# 压缩包密码匹配 + 改名(轻量版)
# 用法:
#   1) 把压缩包/文件夹拖到 MatchRename.bat 上;没匹配到密码会提示输入,验证通过自动记入密码库
#   2) 命令行: powershell -NoProfile -File MatchRename.ps1 [-NoPrompt] [-Rename ask|yes|no] <路径> [<路径>...]
# 密码库: 同目录 passwords.csv (md5,password,note;同一 MD5 可多行 = 多个候选密码)
# 历史:   同目录 history.csv (自动追加)
[CmdletBinding(PositionalBinding = $false)]
param(
  [switch]$NoPrompt,
  [ValidateSet('ask', 'yes', 'no')]
  [string]$Rename = 'ask',
  [Parameter(ValueFromRemainingArguments = $true)]
  [string[]]$Paths
)

$ErrorActionPreference = 'Stop'

# ---------- 配置 ----------
# 7z.exe 查找顺序: 同目录 7z-path.local.txt 第一行写的路径 → PATH → 常见安装位置
$SevenZipLocalConfig = Join-Path $PSScriptRoot '7z-path.local.txt'
$SevenZipCandidates = @(
  'C:\Program Files\7-Zip\7z.exe',
  'C:\Program Files (x86)\7-Zip\7z.exe'
)
$RenameTemplate = '{name}-{password}{ext}'
$ScanExtensions = @('zip', '7z', 'rar', 'tar', 'gz', 'bz2', 'xz', 'tgz', 'iso', 'cab')
$PasswordsCsv   = Join-Path $PSScriptRoot 'passwords.csv'
$HistoryCsv     = Join-Path $PSScriptRoot 'history.csv'

# ---------- 分卷/扩展名识别 ----------

# 识别分卷命名,返回 @{IsFirst; IsPart; BaseKey; BaseName} 或 $null(非压缩包/未知)
function Get-VolumeInfo([string]$fileName) {
  $lower = $fileName.ToLower()
  $m = [regex]::Match($lower, '^(.+)\.part(\d+)\.rar$')
  if ($m.Success) {
    $idx = [int]$m.Groups[2].Value
    return @{ IsFirst = ($idx -eq 1); IsPart = ($idx -gt 1); BaseKey = ($m.Groups[1].Value + '.partN.rar'); BaseName = $fileName.Substring(0, $m.Groups[1].Length) }
  }
  $m = [regex]::Match($lower, '^(.+)\.z(\d{2,})$')
  if ($m.Success) {
    return @{ IsFirst = $false; IsPart = $true; BaseKey = ($m.Groups[1].Value + '.zip'); BaseName = $fileName.Substring(0, $m.Groups[1].Length) }
  }
  $m = [regex]::Match($lower, '^(.+)\.r(\d{2,})$')
  if ($m.Success) {
    return @{ IsFirst = $false; IsPart = $true; BaseKey = ($m.Groups[1].Value + '.rar'); BaseName = $fileName.Substring(0, $m.Groups[1].Length) }
  }
  $m = [regex]::Match($lower, '^(.+)\.(\d{3,})$')
  if ($m.Success) {
    $idx = [int]$m.Groups[2].Value
    return @{ IsFirst = ($idx -eq 1); IsPart = ($idx -gt 1); BaseKey = $m.Groups[1].Value; BaseName = $fileName.Substring(0, $m.Groups[1].Length) }
  }
  $m = [regex]::Match($lower, '^(.+)\.rar$')
  if ($m.Success) {
    return @{ IsFirst = $true; IsPart = $false; BaseKey = ($m.Groups[1].Value + '.rar'); BaseName = $fileName.Substring(0, $m.Groups[1].Length) }
  }
  $m = [regex]::Match($lower, '^(.+)\.zip$')
  if ($m.Success) {
    return @{ IsFirst = $true; IsPart = $false; BaseKey = ($m.Groups[1].Value + '.zip'); BaseName = $fileName.Substring(0, $m.Groups[1].Length) }
  }
  return $null
}

# 多段扩展感知拆名: a.7z.001 -> name=a ext=.7z.001; a.part1.rar / a.tar.gz 同理
function Split-NameExt([string]$base) {
  $m = [regex]::Match($base, '(\.part\d+\.rar)$', 'IgnoreCase')
  if ($m.Success) { return @{ Name = $base.Substring(0, $base.Length - $m.Length); Ext = $base.Substring($base.Length - $m.Length) } }
  $m = [regex]::Match($base, '(\.[^.]+\.\d{3,})$')
  if ($m.Success) { return @{ Name = $base.Substring(0, $base.Length - $m.Length); Ext = $base.Substring($base.Length - $m.Length) } }
  foreach ($e in @('.tar.gz', '.tar.bz2', '.tar.xz')) {
    if ($base.ToLower().EndsWith($e)) { return @{ Name = $base.Substring(0, $base.Length - $e.Length); Ext = $base.Substring($base.Length - $e.Length) } }
  }
  $i = $base.LastIndexOf('.')
  if ($i -le 0) { return @{ Name = $base; Ext = '' } }
  return @{ Name = $base.Substring(0, $i); Ext = $base.Substring($i) }
}

function Test-ArchiveName([string]$name, [hashtable]$extSet) {
  if (Get-VolumeInfo $name) { return $true }
  $parts = $name.ToLower().Split('.')
  if ($parts.Count -lt 2) { return $false }
  if ($extSet.ContainsKey($parts[-1])) { return $true }
  if ($parts.Count -ge 3 -and $extSet.ContainsKey($parts[-2] + '.' + $parts[-1])) { return $true }
  return $false
}

# 返回同目录下同一分卷集的所有卷(含自身);单卷返回自身
function Get-VolumeSiblings([string]$firstPath) {
  $dir = Split-Path -Parent $firstPath
  $name = Split-Path -Leaf $firstPath
  $info = Get-VolumeInfo $name
  if (-not $info) { return @($firstPath) }
  $sibs = @()
  foreach ($f in (Get-ChildItem -LiteralPath $dir -File)) {
    $i2 = Get-VolumeInfo $f.Name
    if ($i2 -and $i2.BaseKey -eq $info.BaseKey) { $sibs += $f.FullName }
  }
  if ($sibs.Count -eq 0) { return @($firstPath) }
  return @($sibs | Sort-Object)
}

# ---------- 密码库 / 历史 ----------

function Get-PasswordMap {
  if (-not (Test-Path -LiteralPath $PasswordsCsv)) {
    [IO.File]::WriteAllText($PasswordsCsv, "md5,password,note`r`n", (New-Object System.Text.UTF8Encoding($true)))
    Write-Host "密码库不存在,已创建空文件: $PasswordsCsv" -ForegroundColor Yellow
    Write-Host "请往里面填 md5,password,note 记录后重试" -ForegroundColor Yellow
    return @{}
  }
  $rows = @(Import-Csv -LiteralPath $PasswordsCsv -Header 'md5', 'password', 'note' -Encoding UTF8)
  $map = @{}
  $start = 0
  if ($rows.Count -gt 0 -and $rows[0].md5 -eq 'md5') { $start = 1 }
  for ($i = $start; $i -lt $rows.Count; $i++) {
    $md5 = ('' + $rows[$i].md5).Trim().ToLower()
    $pwd = '' + $rows[$i].password
    if ($md5 -notmatch '^[0-9a-f]{32}$') { continue }
    if ($pwd -eq '') { continue }
    if (-not $map.ContainsKey($md5)) { $map[$md5] = @() }
    $map[$md5] += $pwd
  }
  return $map
}

function ConvertTo-CsvField([string]$s) {
  if ($null -eq $s) { $s = '' }
  return '"' + $s.Replace('"', '""') + '"'
}

function Add-History([string]$status, [string]$file, [string]$md5, [string]$password, [string]$newFile, [string]$detail) {
  if (-not (Test-Path -LiteralPath $HistoryCsv)) {
    [IO.File]::WriteAllText($HistoryCsv, "time,status,file,md5,password,new_file,detail`r`n", (New-Object System.Text.UTF8Encoding($true)))
  }
  $fields = @((Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $status, $file, $md5, $password, $newFile, $detail)
  $line = ($fields | ForEach-Object { ConvertTo-CsvField $_ }) -join ','
  [IO.File]::AppendAllText($HistoryCsv, $line + "`r`n", (New-Object System.Text.UTF8Encoding($false)))
}

# 把新验证通过的密码追加进密码库(用户手编的文件可能没有结尾换行,先补)
function Add-PasswordRecord([string]$md5, [string]$password) {
  if (-not (Test-Path -LiteralPath $PasswordsCsv)) {
    [IO.File]::WriteAllText($PasswordsCsv, "md5,password,note`r`n", (New-Object System.Text.UTF8Encoding($true)))
  }
  $existing = [IO.File]::ReadAllText($PasswordsCsv)
  $prefix = ''
  if ($existing.Length -gt 0 -and -not $existing.EndsWith("`n")) { $prefix = "`r`n" }
  $note = 'manual ' + (Get-Date -Format 'yyyy-MM-dd')
  $line = $md5 + ',' + (ConvertTo-CsvField $password) + ',' + (ConvertTo-CsvField $note)
  [IO.File]::AppendAllText($PasswordsCsv, $prefix + $line + "`r`n", (New-Object System.Text.UTF8Encoding($false)))
}

# 改名与否本次运行只问一次;-Rename yes/no 或 -NoPrompt 时不问
function Get-RenameDecision {
  if ($null -ne $script:RenameDecision) { return $script:RenameDecision }
  if ($Rename -eq 'yes') { $script:RenameDecision = $true }
  elseif ($Rename -eq 'no') { $script:RenameDecision = $false }
  elseif ($NoPrompt) { $script:RenameDecision = $true }
  else {
    $a = Read-Host '要把密码追加到文件名吗? [Y=改名 / n=只记录密码] (本次拖入统一生效)'
    $script:RenameDecision = ($null -eq $a -or $a -eq '' -or $a -match '^[yY]')
  }
  return $script:RenameDecision
}

# ---------- 验证 / 改名 ----------

function Test-ArchivePassword([string]$file, [string]$password) {
  & $script:SevenZip t $file "-p$password" -y -bso0 -bsp0 -bse0 | Out-Null
  return ($LASTEXITCODE -eq 0)
}

function Get-NewName([string]$fileName, [string]$password, [string]$md5) {
  $split = Split-NameExt $fileName
  $safe = $password -replace '[\\/:*?"<>|]', '_'
  return $RenameTemplate.Replace('{name}', $split.Name).Replace('{password}', $safe).Replace('{ext}', $split.Ext).Replace('{md5}', $md5)
}

function Get-AvailableTarget([string]$dir, [string]$desiredName) {
  $target = Join-Path $dir $desiredName
  if (-not (Test-Path -LiteralPath $target)) { return $target }
  $split = Split-NameExt $desiredName
  for ($i = 1; $i -le 9999; $i++) {
    $cand = Join-Path $dir ($split.Name + '_' + $i + $split.Ext)
    if (-not (Test-Path -LiteralPath $cand)) { return $cand }
  }
  throw 'rename collision overflow'
}

function Undo-Renames($done) {
  for ($i = $done.Count - 1; $i -ge 0; $i--) {
    try { Rename-Item -LiteralPath $done[$i].New -NewName (Split-Path -Leaf $done[$i].Old) } catch { }
  }
}

# 分卷整组改名,事务式:任一卷失败则回滚已改的卷。返回 @{Old;New} 列表(首卷在第 0 位)
function Invoke-VolumeSetRename([string[]]$parts, [string]$newFirstPath) {
  $firstOld = $parts[0]
  $dir = Split-Path -Parent $firstOld
  $target = Get-AvailableTarget $dir (Split-Path -Leaf $newFirstPath)
  if ($firstOld -eq $target) { return @(@{ Old = $firstOld; New = $target }) }
  Rename-Item -LiteralPath $firstOld -NewName (Split-Path -Leaf $target)
  $done = New-Object System.Collections.ArrayList
  [void]$done.Add(@{ Old = $firstOld; New = $target })
  if ($parts.Count -eq 1) { return $done }
  $oldBase = (Split-NameExt (Split-Path -Leaf $firstOld)).Name
  $newBase = (Split-NameExt (Split-Path -Leaf $target)).Name
  foreach ($part in ($parts | Select-Object -Skip 1)) {
    $partName = Split-Path -Leaf $part
    if (-not $partName.ToLower().StartsWith($oldBase.ToLower() + '.')) { continue }
    $suffix = $partName.Substring($oldBase.Length)
    $newPartPath = Join-Path $dir ($newBase + $suffix)
    if (Test-Path -LiteralPath $newPartPath) {
      Undo-Renames $done
      throw ('分卷改名冲突: ' + $newBase + $suffix + ' 已存在')
    }
    try {
      Rename-Item -LiteralPath $part -NewName ($newBase + $suffix)
      [void]$done.Add(@{ Old = $part; New = $newPartPath })
    } catch {
      Undo-Renames $done
      throw ('分卷改名失败 (' + $partName + '): ' + $_.Exception.Message)
    }
  }
  return $done
}

# ---------- 输入收集 ----------

# 从 bat 传来的原始 CMDCMDLINE 中解析拖拽的路径(抗 & ! 空格)
function Get-DraggedPaths([string]$cmdLine, [string]$batPath) {
  $rest = $null
  if ($batPath) {
    $idx = $cmdLine.IndexOf($batPath, [System.StringComparison]::OrdinalIgnoreCase)
    if ($idx -ge 0) { $rest = $cmdLine.Substring($idx + $batPath.Length) }
  }
  if ($null -eq $rest) {
    $m = [regex]::Match($cmdLine, '\.bat"?', 'IgnoreCase')
    if (-not $m.Success) { return @() }
    $rest = $cmdLine.Substring($m.Index + $m.Length)
  }
  # 去掉紧跟 bat 路径的闭引号(cmd 的 ""bat" args" 包裹)
  $rest = $rest.TrimStart('"')
  $tokens = @()
  $cur = New-Object System.Text.StringBuilder
  $inQ = $false
  foreach ($ch in $rest.ToCharArray()) {
    if ($ch -eq '"') { $inQ = -not $inQ; continue }
    if (-not $inQ -and ($ch -eq ' ' -or $ch -eq "`t")) {
      if ($cur.Length -gt 0) { $tokens += $cur.ToString(); [void]$cur.Clear() }
      continue
    }
    [void]$cur.Append($ch)
  }
  if ($cur.Length -gt 0) { $tokens += $cur.ToString() }
  return $tokens
}

# 目录递归展开 + 去重 + 分卷分组(每个分卷集只保留入口卷)
function Get-TargetFiles([string[]]$inputs) {
  $extSet = @{}
  foreach ($e in $ScanExtensions) { $extSet[$e] = $true }
  $seen = @{}
  $files = @()
  foreach ($p in $inputs) {
    if (-not (Test-Path -LiteralPath $p)) {
      Write-Host "找不到路径,跳过: $p" -ForegroundColor Yellow
      continue
    }
    $item = Get-Item -LiteralPath $p
    if ($item.PSIsContainer) {
      foreach ($f in (Get-ChildItem -LiteralPath $p -Recurse -File)) {
        if ((Test-ArchiveName $f.Name $extSet) -and -not $seen.ContainsKey($f.FullName.ToLower())) {
          $seen[$f.FullName.ToLower()] = $true
          $files += $f
        }
      }
    } else {
      if (-not $seen.ContainsKey($item.FullName.ToLower())) {
        $seen[$item.FullName.ToLower()] = $true
        $files += $item
      }
    }
  }
  $standalone = @()
  $groups = @{}
  $order = @()
  foreach ($f in $files) {
    $info = Get-VolumeInfo $f.Name
    if (-not $info) { $standalone += $f; continue }
    $key = (Split-Path -Parent $f.FullName).ToLower() + '|' + $info.BaseKey
    if (-not $groups.ContainsKey($key)) { $groups[$key] = @(); $order += $key }
    $groups[$key] += $f
  }
  $out = @() + $standalone
  foreach ($key in $order) {
    $arr = @($groups[$key] | Sort-Object Name)
    $first = $null
    foreach ($f in $arr) {
      if ((Get-VolumeInfo $f.Name).IsFirst) { $first = $f; break }
    }
    if (-not $first) { $first = $arr[0] }
    $out += $first
  }
  return $out
}

# ---------- 主流程 ----------

$inputPaths = @()
if ($env:APM_CMDLINE) { $inputPaths = @(Get-DraggedPaths $env:APM_CMDLINE $env:APM_BAT) }
if ($inputPaths.Count -eq 0 -and $Paths) { $inputPaths = $Paths }

if ($inputPaths.Count -eq 0) {
  Write-Host '压缩包密码匹配 + 改名'
  Write-Host ''
  Write-Host '用法: 把压缩包或文件夹拖到 MatchRename.bat 上'
  Write-Host '      没匹配到密码时会提示输入,验证通过自动记入密码库,改名前会询问'
  Write-Host '      或: powershell -NoProfile -File MatchRename.ps1 [-NoPrompt] [-Rename ask|yes|no] <路径> [<路径>...]'
  Write-Host ''
  Write-Host "密码库: $PasswordsCsv"
  Write-Host '格式:   md5,password,note(同一 MD5 可多行 = 多个候选密码)'
  exit 0
}

$SevenZip = $null
if (Test-Path -LiteralPath $SevenZipLocalConfig) {
  $localPath = ('' + (Get-Content -LiteralPath $SevenZipLocalConfig -TotalCount 1)).Trim()
  if ($localPath -and (Test-Path -LiteralPath $localPath)) { $SevenZip = $localPath }
}
if (-not $SevenZip) {
  $cmd = Get-Command 7z.exe -ErrorAction SilentlyContinue
  if ($cmd) { $SevenZip = $cmd.Source }
}
if (-not $SevenZip) {
  foreach ($c in $SevenZipCandidates) {
    if (Test-Path -LiteralPath $c) { $SevenZip = $c; break }
  }
}
if (-not $SevenZip) {
  Write-Host '找不到 7z.exe,三种解决方式任选:' -ForegroundColor Red
  Write-Host '  1) 安装 7-Zip 到默认位置 (Program Files)'
  Write-Host '  2) 把 7-Zip 目录加入 PATH'
  Write-Host '  3) 在脚本同目录建 7z-path.local.txt,第一行写 7z.exe 完整路径'
  exit 1
}

$map = Get-PasswordMap
$md5Count = $map.Keys.Count
$recCount = 0
foreach ($k in $map.Keys) { $recCount += $map[$k].Count }
Write-Host "密码库: $recCount 条记录 / $md5Count 个 MD5" -ForegroundColor Cyan

$targets = @(Get-TargetFiles $inputPaths)
if ($targets.Count -eq 0) {
  Write-Host '没有找到压缩包文件' -ForegroundColor Yellow
  exit 0
}
Write-Host "待处理: $($targets.Count) 个压缩包"
Write-Host ''

$countOk = 0
$countRecorded = 0
$countFail = 0
$countNone = 0
$countErr = 0
$script:RenameDecision = $null
$script:SessionPasswords = New-Object System.Collections.ArrayList

foreach ($f in $targets) {
  $md5 = $null
  try {
    $md5 = (Get-FileHash -LiteralPath $f.FullName -Algorithm MD5).Hash.ToLower()
  } catch {
    Write-Host "[错误] $($f.Name) — 计算 MD5 失败: $($_.Exception.Message)" -ForegroundColor Red
    Add-History 'error' $f.FullName '' '' '' ('md5 failed: ' + $_.Exception.Message)
    $countErr++
    continue
  }

  $cands = @()
  if ($map.ContainsKey($md5)) { $cands = @($map[$md5]) }

  # 先试库里的候选,再试本次运行里输过的密码,最后交互式问用户
  $matched = $null
  $fromLibrary = $false
  foreach ($pwd in $cands) {
    if (Test-ArchivePassword $f.FullName $pwd) { $matched = $pwd; $fromLibrary = $true; break }
  }
  if ($null -eq $matched) {
    foreach ($pwd in $script:SessionPasswords) {
      if ($cands -contains $pwd) { continue }
      if (Test-ArchivePassword $f.FullName $pwd) { $matched = $pwd; break }
    }
  }
  $prompted = $false
  if ($null -eq $matched -and -not $NoPrompt) {
    $prompted = $true
    if ($cands.Count -gt 0) {
      Write-Host "[?] $($f.Name) — 库里 $($cands.Count) 个候选都没验证通过" -ForegroundColor Yellow
    } else {
      Write-Host "[?] $($f.Name) — 密码库无记录  md5=$md5" -ForegroundColor Yellow
    }
    while ($true) {
      $inp = Read-Host '    输入密码(直接回车跳过)'
      if ($null -eq $inp -or $inp -eq '') { break }
      if (Test-ArchivePassword $f.FullName $inp) { $matched = $inp; break }
      Write-Host '    密码不对,可以再试' -ForegroundColor Red
    }
  }

  if ($null -eq $matched) {
    if ($cands.Count -gt 0) {
      if (-not $prompted) { Write-Host "[失败] $($f.Name) — $($cands.Count) 个候选密码都没验证通过" -ForegroundColor Red }
      else { Write-Host "    已跳过" -ForegroundColor DarkYellow }
      Add-History 'fail' $f.FullName $md5 ($cands -join ' | ') '' 'no candidate verified'
      $countFail++
    } else {
      if (-not $prompted) { Write-Host "[无记录] $($f.Name)  md5=$md5" -ForegroundColor Yellow }
      else { Write-Host "    已跳过" -ForegroundColor DarkYellow }
      Add-History 'none' $f.FullName $md5 '' '' 'no password record'
      $countNone++
    }
    continue
  }

  if (-not $script:SessionPasswords.Contains($matched)) { [void]$script:SessionPasswords.Add($matched) }
  if (-not $fromLibrary) {
    Add-PasswordRecord $md5 $matched
    Write-Host "[记录] $($f.Name) 的密码已写入密码库" -ForegroundColor Cyan
  }

  $safePwd = $matched -replace '[\\/:*?"<>|]', '_'
  $split = Split-NameExt $f.Name
  if ($split.Name.EndsWith('-' + $safePwd)) {
    Write-Host "[已命名] $($f.Name)  密码=$matched(文件名里已带,跳过改名)" -ForegroundColor Green
    Add-History 'ok' $f.FullName $md5 $matched $f.FullName 'already named'
    $countOk++
    continue
  }

  if (-not (Get-RenameDecision)) {
    Write-Host "[仅记录] $($f.Name)  密码=$matched(按选择不改名)" -ForegroundColor Green
    Add-History 'recorded' $f.FullName $md5 $matched '' 'rename skipped by choice'
    $countRecorded++
    continue
  }

  $sibs = @(Get-VolumeSiblings $f.FullName)
  $parts = @($f.FullName) + @($sibs | Where-Object { $_ -ne $f.FullName })
  $newName = Get-NewName $f.Name $matched $md5
  try {
    $done = @(Invoke-VolumeSetRename $parts (Join-Path $f.DirectoryName $newName))
    $newFirst = Split-Path -Leaf $done[0].New
    $volNote = ''
    if ($done.Count -gt 1) { $volNote = "(分卷 x$($done.Count) 整组改名)" }
    Write-Host "[成功] $($f.Name) -> $newFirst  密码=$matched $volNote" -ForegroundColor Green
    Add-History 'ok' $f.FullName $md5 $matched $done[0].New $volNote
    $countOk++
  } catch {
    Write-Host "[错误] $($f.Name) — 改名失败: $($_.Exception.Message)" -ForegroundColor Red
    Add-History 'error' $f.FullName $md5 $matched '' ('rename failed: ' + $_.Exception.Message)
    $countErr++
  }
}

Write-Host ''
Write-Host ('完成: 改名 ' + $countOk + ' | 仅记录 ' + $countRecorded + ' | 失败 ' + $countFail + ' | 无记录 ' + $countNone + ' | 错误 ' + $countErr) -ForegroundColor Cyan
Write-Host "历史: $HistoryCsv"
exit 0
