<#
.SYNOPSIS
    Скачивание работ выбранных авторов с e6ai.net через публичный API (как у e621).

.DESCRIPTION
    - Авторы задаются параметром -Authors, файлом -AuthorsFile или вводом с клавиатуры.
    - Сначала собирается список постов и показывается сводка, потом спрашивается подтверждение.
    - Уже скачанные файлы пропускаются (проверка по имени и размеру), поэтому скрипт можно
      запускать повторно: он докачает только новое и недокачанное.
    - Каждый файл проверяется по размеру и MD5 из API. Битые файлы не сохраняются.
    - Существующие файлы не перезаписываются, если они совпадают по размеру.

.EXAMPLE
    .\download_e6ai.ps1 -Authors towersmash,namsi -DryRun
    Только посчитать, что и сколько будет скачано. Ничего не скачивается.

.EXAMPLE
    .\download_e6ai.ps1 -Authors towersmash,namsi
    Показать сводку, спросить подтверждение и скачать.

.EXAMPLE
    .\download_e6ai.ps1 -AuthorsFile .\authors.txt -OutDir D:\Archive\e6ai -Yes
    Авторы из файла (по одному в строке, # - комментарий), без вопроса подтверждения.

.NOTES
    Правила API: не больше 2 запросов в секунду (скрипт делает паузу, по умолчанию 600 мс),
    обязателен осмысленный User-Agent. Лучше указать свой ник: -UserAgent "MyArchive/1.0 (by ВАШ_НИК)".
    Запуск при блокировке скриптов:  powershell -ExecutionPolicy Bypass -File .\download_e6ai.ps1 -Authors namsi
#>
[CmdletBinding()]
param(
    [string[]]$Authors,
    [string]$AuthorsFile,
    [string]$OutDir,
    [string]$UserAgent = "PersonalArchive/1.0",
    [switch]$DryRun,
    [switch]$Yes,
    [int]$DelayMs = 600,
    [int]$PageSize = 320,
    [int]$MaxRetries = 4,
    [string]$BaseUrl = "https://e6ai.net"
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"   # в Windows PowerShell 5.1 прогресс-бар сильно тормозит скачивание
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# ---------- настройки и проверки ----------
$BaseUrl = $BaseUrl.TrimEnd("/")
$isRealHost = $BaseUrl -match '^https://e6ai\.net$'
if ($isRealHost -and $DelayMs -lt 500) {
    Write-Host "Пауза меньше 500 мс нарушает лимит API (2 запроса/с), ставлю 500 мс." -ForegroundColor Yellow
    $DelayMs = 500
}
if ($PageSize -lt 1) { $PageSize = 1 }
if ($PageSize -gt 320) { $PageSize = 320 }   # максимум API
if ([string]::IsNullOrWhiteSpace($UserAgent)) {
    Write-Host "[ОШИБКА] User-Agent не должен быть пустым." -ForegroundColor Red
    exit 1
}

if (-not $OutDir) {
    $root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $OutDir = Join-Path $root "e6ai_archive"
}
try {
    New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
    $OutDir = (Resolve-Path -LiteralPath $OutDir).Path
} catch {
    Write-Host "[ОШИБКА] Не удалось создать или открыть папку: $OutDir" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

$script:Headers = @{ "User-Agent" = $UserAgent }
$LogFile    = Join-Path $OutDir "download_log.txt"
$FailedFile = Join-Path $OutDir "failed.txt"

# ---------- вспомогательные функции ----------
function Write-Log([string]$Message, [string]$Color) {
    Add-Content -LiteralPath $LogFile -Encoding UTF8 -Value ("{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message)
    if ($Color) { Write-Host $Message -ForegroundColor $Color } else { Write-Host $Message }
}

function Get-HttpCode($ErrorRecord) {
    try { return [int]$ErrorRecord.Exception.Response.StatusCode } catch { return 0 }
}

# Повторяет запрос при сетевых ошибках, 429 и 5xx. Остальные ошибки (403, 404...) пробрасывает сразу.
function Invoke-Retry([scriptblock]$Action, [string]$What) {
    $wait = 3
    for ($try = 1; $try -le $MaxRetries; $try++) {
        try { return (& $Action) }
        catch {
            $code = Get-HttpCode $_
            $retryable = ($code -eq 0) -or ($code -eq 429) -or ($code -ge 500)
            if (-not $retryable -or $try -ge $MaxRetries) { throw }
            Write-Log "  повтор $try/$MaxRetries ($What): $($_.Exception.Message) [HTTP $code], жду $wait c" "DarkYellow"
            Start-Sleep -Seconds $wait
            $wait = $wait * 2
        }
    }
}

function ConvertTo-Tag([string]$Name) { return ($Name.Trim().ToLower() -replace '\s+', '_') }
function ConvertTo-SafeName([string]$Name) { return (($Name -replace '[\\/:*?"<>|]', '_').Trim('. ')) }
function Format-Size([long]$Bytes) {
    if ($Bytes -ge 1GB) { return ("{0:N2} GB" -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ("{0:N1} MB" -f ($Bytes / 1MB)) }
    return ("{0:N0} KB" -f [math]::Ceiling($Bytes / 1KB))
}

function Get-AuthorPosts([string]$Tag) {
    $seen = @{}
    $list = New-Object System.Collections.Generic.List[object]
    $page = $null
    $prevLast = [long]::MaxValue
    while ($true) {
        $url = "$BaseUrl/posts.json?tags=" + [uri]::EscapeDataString($Tag) + "&limit=$PageSize"
        if ($page) { $url += "&page=$page" }
        # Invoke-WebRequest + ConvertFrom-Json: у Invoke-RestMethod в Windows PowerShell 5.1 лимит ответа 2 МБ
        $resp  = Invoke-Retry { (Invoke-WebRequest -Uri $url -Headers $script:Headers -UseBasicParsing -TimeoutSec 60).Content | ConvertFrom-Json } "список '$Tag'"
        $posts = @($resp.posts)
        if ($posts.Count -eq 0) { break }
        foreach ($p in $posts) {
            $key = [string]$p.id
            if (-not $seen.ContainsKey($key)) { $seen[$key] = $true; $list.Add($p) }
        }
        $lastId = [long](($posts | Measure-Object -Property id -Minimum).Minimum)
        Write-Host ("  получено постов: {0}, последний ID: {1}" -f $list.Count, $lastId)
        if ($lastId -ge $prevLast) { break }   # защита от зацикливания
        $prevLast = $lastId
        $page = "b$lastId"
        Start-Sleep -Milliseconds $DelayMs
    }
    return $list.ToArray()
}

function ConvertTo-Item($Post, [string]$Dir) {
    $id  = [long]$Post.id
    $ext = ([string]$Post.file.ext).ToLower()
    if ($ext -notmatch '^[a-z0-9]{1,5}$') { $ext = "bin" }
    [pscustomobject]@{
        Id       = $id
        Ext      = $ext
        Url      = [string]$Post.file.url
        Size     = [long]$Post.file.size
        Md5      = ([string]$Post.file.md5).ToLower()
        Uploader = [string]$Post.uploader_name
        Path     = (Join-Path $Dir "$id.$ext")
    }
}

function Save-File($Item) {
    $part = "$($Item.Path).part"
    Invoke-Retry { Invoke-WebRequest -Uri $Item.Url -Headers $script:Headers -OutFile $part -UseBasicParsing -TimeoutSec 600 } "файл $($Item.Id)" | Out-Null
    $len = (Get-Item -LiteralPath $part).Length
    if ($Item.Size -gt 0 -and $len -ne $Item.Size) {
        Remove-Item -LiteralPath $part -Force
        throw "размер $len байт, ожидалось $($Item.Size)"
    }
    if ($Item.Md5 -match '^[0-9a-f]{32}$') {
        $hash = (Get-FileHash -LiteralPath $part -Algorithm MD5).Hash.ToLower()
        if ($hash -ne $Item.Md5) {
            Remove-Item -LiteralPath $part -Force
            throw "MD5 не совпал (получен $hash, ожидался $($Item.Md5))"
        }
    }
    Move-Item -LiteralPath $part -Destination $Item.Path -Force
}

# ---------- список авторов ----------
$names = @()
if ($Authors)     { $names += $Authors }
if ($AuthorsFile) {
    if (-not (Test-Path -LiteralPath $AuthorsFile)) {
        Write-Host "[ОШИБКА] Файл с авторами не найден: $AuthorsFile" -ForegroundColor Red
        exit 1
    }
    $names += Get-Content -LiteralPath $AuthorsFile -Encoding UTF8 | Where-Object { $_.Trim() -ne "" -and -not $_.Trim().StartsWith("#") }
}
if ($names.Count -eq 0) {
    $raw = Read-Host "Авторы (теги художников) через запятую"
    $names += $raw
}
# при запуске через -File массив приходит одной строкой "a,b", поэтому режем по запятой
$tags = @($names | ForEach-Object { $_ -split ',' } | ForEach-Object { ConvertTo-Tag $_ } | Where-Object { $_ } | Select-Object -Unique)
if ($tags.Count -eq 0) {
    Write-Host "[ОШИБКА] Не указано ни одного автора." -ForegroundColor Red
    exit 1
}
foreach ($t in $tags) {
    if ($t -match '^[-~]' -or $t -match '[:*,]') {
        Write-Host "[ОШИБКА] Недопустимое имя тега: '$t' (без -, ~, :, *). Нужно обычное имя автора." -ForegroundColor Red
        exit 1
    }
}

Write-Log ("=== Запуск: авторы: {0}; папка: {1}; режим: {2} ===" -f ($tags -join ", "), $OutDir, $(if ($DryRun) { "DryRun" } else { "загрузка" }))

# ---------- этап 1: собираем списки ----------
$plan = @()
$totalTodo = 0
$totalBytes = [long]0
foreach ($tag in $tags) {
    Write-Host ""
    Write-Host "========== $tag ==========" -ForegroundColor Cyan
    $dir = Join-Path $OutDir (ConvertTo-SafeName $tag)
    try {
        $posts = @(Get-AuthorPosts $tag)
    } catch {
        Write-Log "[ОШИБКА] Не удалось получить список для '$tag': $($_.Exception.Message)" "Red"
        continue
    }
    if ($posts.Count -eq 0) {
        Write-Log "Постов по тегу '$tag' не найдено. Проверь написание (нужен тег художника, как на сайте)." "Yellow"
        continue
    }

    $items   = @($posts | ForEach-Object { ConvertTo-Item $_ $dir })
    $noUrl   = @($items | Where-Object { $_.Url -notmatch '^https?://' })
    $withUrl = @($items | Where-Object { $_.Url -match '^https?://' })
    $have    = @($withUrl | Where-Object { (Test-Path -LiteralPath $_.Path) -and ($_.Size -le 0 -or (Get-Item -LiteralPath $_.Path).Length -eq $_.Size) })
    $haveIds = @{}; foreach ($h in $have) { $haveIds[$h.Id] = $true }
    $todo    = @($withUrl | Where-Object { -not $haveIds.ContainsKey($_.Id) })
    $todoBytes = [long](($todo | Measure-Object -Property Size -Sum).Sum)

    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $items | Select-Object Id, Ext, Size, Md5, Uploader, Url |
        Export-Csv -LiteralPath (Join-Path $dir "_posts.csv") -NoTypeInformation -Encoding UTF8

    Write-Log ("Всего постов: {0} | без ссылки: {1} | уже скачано: {2} | к загрузке: {3} ({4})" -f $items.Count, $noUrl.Count, $have.Count, $todo.Count, (Format-Size $todoBytes)) "Green"
    if ($noUrl.Count -gt 0) {
        Write-Host "  У $($noUrl.Count) постов API не отдал ссылку на файл (возможно, скрыты настройками сайта или нужен вход). Они пропущены." -ForegroundColor Yellow
    }
    $todo | Group-Object Ext | Sort-Object Count -Descending | ForEach-Object { Write-Host ("    {0,-6} {1}" -f $_.Name, $_.Count) }

    $plan += [pscustomobject]@{ Tag = $tag; Dir = $dir; Todo = $todo }
    $totalTodo  += $todo.Count
    $totalBytes += $todoBytes
}

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ("Итого к загрузке: {0} файлов, {1}" -f $totalTodo, (Format-Size $totalBytes)) -ForegroundColor Cyan
Write-Host "Папка: $OutDir"

if ($totalTodo -eq 0) { Write-Log "Скачивать нечего." "Green"; exit 0 }
if ($DryRun) { Write-Log "DryRun: ничего не скачано. Списки сохранены в _posts.csv каждого автора." "Yellow"; exit 0 }

if (-not $Yes) {
    $ans = Read-Host "Скачать? (y/n)"
    if ($ans.Trim().ToLower() -notmatch '^(y|yes|д|да)$') { Write-Log "Отменено пользователем." "Yellow"; exit 0 }
}

# ---------- этап 2: загрузка ----------
$ok = 0; $failed = 0
foreach ($entry in $plan) {
    Write-Host ""
    Write-Host "---------- загрузка: $($entry.Tag) ----------" -ForegroundColor Cyan
    $n = 0
    foreach ($item in $entry.Todo) {
        $n++
        $label = "[{0}/{1}] {2}.{3}" -f $n, $entry.Todo.Count, $item.Id, $item.Ext
        try {
            Save-File $item
            $ok++
            Write-Log "$label  OK ($(Format-Size $item.Size))"
        } catch {
            $failed++
            Write-Log "$label  ОШИБКА: $($_.Exception.Message)" "Red"
            Add-Content -LiteralPath $FailedFile -Encoding UTF8 -Value ("{0}`t{1}`t{2}`t{3}" -f $entry.Tag, $item.Id, $item.Url, $_.Exception.Message)
        }
        Start-Sleep -Milliseconds $DelayMs
    }
}

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Log "ИТОГ: скачано $ok, ошибок $failed" $(if ($failed -gt 0) { "Yellow" } else { "Green" })
if ($failed -gt 0) { Write-Host "Список неудач: $FailedFile. Повторный запуск скрипта докачает только их." -ForegroundColor Yellow }
Write-Host "Лог: $LogFile"