# Booru Downloader - этап 1: проверка поиска (без скачивания)
# Совместимость: Windows PowerShell 5.1 / Windows 11
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$Root = $PSScriptRoot
$ConfigPath = Join-Path $Root 'config.ini'
$ArchiveRoot = Join-Path $Root 'Archive'
$LogRoot = Join-Path $Root 'Logs'
New-Item -ItemType Directory -Force -Path $ArchiveRoot,$LogRoot | Out-Null
$LogFile = Join-Path $LogRoot 'search.log'

function Write-Log([string]$Message, [string]$Color = 'Gray') {
    $line = '{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8
    Write-Host $Message -ForegroundColor $Color
}
function Read-Ini([string]$Path) {
    $data = @{}
    $section = 'General'
    foreach ($line in (Get-Content -LiteralPath $Path -Encoding UTF8)) {
        $s = $line.Trim()
        if (-not $s -or $s.StartsWith(';') -or $s.StartsWith('#')) { continue }
        if ($s -match '^\[(.+)\]$') { $section = $Matches[1]; if (-not $data.ContainsKey($section)) {$data[$section]=@{}}; continue }
        if ($s -match '^([^=]+)=(.*)$') { if (-not $data.ContainsKey($section)) {$data[$section]=@{}}; $data[$section][$Matches[1].Trim()]=$Matches[2].Trim() }
    }
    return $data
}
function Get-Cfg($Cfg,[string]$Section,[string]$Key,[string]$Default='') {
    if ($Cfg.ContainsKey($Section) -and $Cfg[$Section].ContainsKey($Key)) { return [string]$Cfg[$Section][$Key] }
    return $Default
}
function Encode([string]$Value) { return [uri]::EscapeDataString($Value) }
function Invoke-Get([string]$Url,[hashtable]$Headers,[int]$Timeout) {
    return Invoke-RestMethod -Uri $Url -Method Get -Headers $Headers -TimeoutSec $Timeout -UseBasicParsing
}
function Get-CountValue($Object) {
    if ($null -eq $Object) { return $null }
    if ($Object.counts -and $null -ne $Object.counts.posts) { return [long]$Object.counts.posts }
    if ($Object.'@attributes' -and $null -ne $Object.'@attributes'.count) { return [long]$Object.'@attributes'.count }
    if ($Object.Count -and $Object.Count -is [int]) { return [long]$Object.Count }
    return $null
}
function Search-E621([string]$Tag,$Cfg,[hashtable]$Headers,[int]$Timeout,[int]$Limit) {
    $h = @{} + $Headers
    $login = Get-Cfg $Cfg 'e621' 'Login'
    $key = Get-Cfg $Cfg 'e621' 'ApiKey'
    if ($login -and $key) {
        $pair = '{0}:{1}' -f $login,$key
        $h['Authorization'] = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
    }
    $q = Encode $Tag
    $countObj = Invoke-Get "https://e621.net/posts/count.json?tags=$q" $h $Timeout
    $count = $null
    if ($null -ne $countObj -and $null -ne $countObj.counts -and $null -ne $countObj.counts.posts) {
        $count = [long]$countObj.counts.posts
    }

    # /posts.json возвращает объект-обёртку: { "posts": [ ... ] }.
    # Нельзя считать сам объект ответом с одной записью.
    $postsResponse = Invoke-Get "https://e621.net/posts.json?tags=$q&limit=$Limit" $h $Timeout
    $posts = @()
    if ($null -ne $postsResponse -and $null -ne $postsResponse.posts) {
        $posts = @($postsResponse.posts)
    } elseif ($postsResponse -is [System.Array]) {
        $posts = @($postsResponse)
    }

    $note = $null
    if ($null -eq $count) {
        $count = $posts.Count
        $note = 'API не вернул общий счётчик; указано число записей на первой странице'
    }
    return [pscustomobject]@{Count=$count; Sample=$posts; Method='e621 JSON API'; SearchUrl="https://e621.net/posts?tags=$q"; Note=$note}
}
function Search-GelbooruLike([string]$HostName,[string]$Tag,$Cfg,[string]$Section,[hashtable]$Headers,[int]$Timeout,[int]$Limit) {
    $base = "https://$HostName/index.php"
    $query = "page=dapi&s=post&q=index&json=1&limit=$Limit&tags=$(Encode $Tag)"
    if ($Section -eq 'Gelbooru') {
        $uid = Get-Cfg $Cfg 'Gelbooru' 'UserId'; $key = Get-Cfg $Cfg 'Gelbooru' 'ApiKey'
        if ($uid -and $key) { $query += "&user_id=$(Encode $uid)&api_key=$(Encode $key)" }
    }
    if ($Section -eq 'Rule34') {
        $uid = Get-Cfg $Cfg 'Rule34' 'UserId'; $key = Get-Cfg $Cfg 'Rule34' 'ApiKey'
        if ($uid -and $key) { $query += "&user_id=$(Encode $uid)&api_key=$(Encode $key)" }
    }
    $obj = Invoke-Get "$base`?$query" $Headers $Timeout
    $posts = @()
    if ($obj.post) { $posts = @($obj.post) }
    elseif ($obj.posts) { $posts = @($obj.posts) }
    elseif ($obj -is [System.Array]) { $posts = @($obj) }
    $count = Get-CountValue $obj
    if ($null -eq $count) {
        # Некоторые реализации возвращают только первую страницу без общего счётчика.
        $count = $posts.Count
        $countNote = 'общее число API не сообщил; показано число в ответе'
    } else { $countNote = 'общее число из API' }
    return [pscustomobject]@{Count=$count; Sample=$posts; Method="$Section DAPI"; SearchUrl="https://$HostName/index.php?page=post&s=list&tags=$(Encode $Tag)"; Note=$countNote}
}

function Search-Danbooru([string]$HostName,[string]$SiteName,[string]$Tag,[hashtable]$Headers,[int]$Timeout,[int]$Limit) {
    # Blep.art и Floofgens используют Danbooru-подобный API:
    # GET /posts.json?tags=...&limit=...
    # Для обоих сайтов обязателен непустой описательный User-Agent.
    $q = Encode $Tag
    $url = "https://$HostName/posts.json?tags=$q&limit=$Limit"
    $obj = Invoke-Get $url $Headers $Timeout
    $posts = @()
    if ($null -ne $obj -and $null -ne $obj.posts) {
        $posts = @($obj.posts)
    } elseif ($obj -is [System.Array]) {
        $posts = @($obj)
    } elseif ($null -ne $obj -and $obj.id) {
        $posts = @($obj)
    }
    return [pscustomobject]@{
        Count=$posts.Count
        Sample=$posts
        Method="$SiteName JSON API (/posts.json)"
        SearchUrl="https://$HostName/posts?tags=$q"
        Note='Показано число записей в ответе API; общий счётчик отдельно не запрашивался'
    }
}

function Search-Nozomi([string]$Tag,[hashtable]$Headers,[int]$Timeout) {
    # Nozomi использует отдельные бинарные индексы тегов, а не Danbooru JSON.
    $encoded = Encode $Tag
    $indexUrl = "https://j.nozomi.la/nozomi/$encoded.nozomi"
    $request = [Net.HttpWebRequest]::Create($indexUrl)
    $request.Method = 'HEAD'; $request.Timeout = $Timeout * 1000
    if ($Headers.ContainsKey('User-Agent')) { $request.UserAgent = $Headers['User-Agent'] }
    if ($Headers.ContainsKey('Accept')) { $request.Accept = $Headers['Accept'] }
    $bytes = [long]-1
    try {
        $response = $request.GetResponse()
        try { $bytes = [long]$response.ContentLength } finally { $response.Close() }
    } catch {
        # Некоторые CDN не поддерживают HEAD. Ниже будет попытка Range GET.
    }
    if ($bytes -le 0) {
        $rangeRequest = [Net.HttpWebRequest]::Create($indexUrl)
        $rangeRequest.Method = 'GET'; $rangeRequest.Timeout = $Timeout * 1000
        $rangeRequest.AddRange(0, 0)
        if ($Headers.ContainsKey('User-Agent')) { $rangeRequest.UserAgent = $Headers['User-Agent'] }
        $rangeResponse = $rangeRequest.GetResponse()
        try {
            $contentRange = [string]$rangeResponse.Headers['Content-Range']
            if ($contentRange -match '/(\d+)$') { $bytes = [long]$Matches[1] }
        } finally { $rangeResponse.Close() }
    }
    if ($bytes -le 0) { throw 'Сервер не сообщил размер бинарного индекса (HEAD/Range). Точное количество определить нельзя.' }
    $approx = [long][math]::Floor($bytes / 4)
    return [pscustomobject]@{Count=$approx; Sample=@(); Method='Nozomi tag index (HEAD)'; SearchUrl="https://nozomi.la/search.html?q=$(Encode $Tag)"; Note='оценка по размеру бинарного индекса; требуется сверка с сайтом'}
}

if (-not (Test-Path -LiteralPath $ConfigPath)) { Write-Host "Не найден config.ini: $ConfigPath" -ForegroundColor Red; exit 1 }
$Cfg = Read-Ini $ConfigPath
$userAgent = Get-Cfg $Cfg 'General' 'UserAgent' 'PersonalBooruArchive/0.1'
$timeout = 45; [void][int]::TryParse((Get-Cfg $Cfg 'General' 'TimeoutSec' '45'),[ref]$timeout)
$limit = 100; [void][int]::TryParse((Get-Cfg $Cfg 'General' 'PageSize' '100'),[ref]$limit)
if ($limit -lt 1) {$limit=1}; if ($limit -gt 320) {$limit=320}
$headers = @{ 'User-Agent' = $userAgent; 'Accept'='application/json, text/html, */*' }

$sites = @(
    [pscustomobject]@{Name='e621'; Key='e621'},
    [pscustomobject]@{Name='Blep.art'; Key='Blep'},
    [pscustomobject]@{Name='Floofgens.net'; Key='Floofgens'},
    [pscustomobject]@{Name='Gelbooru'; Key='Gelbooru'},
    [pscustomobject]@{Name='Rule34.xxx'; Key='Rule34'},
    [pscustomobject]@{Name='TBIB'; Key='TBIB'},
    [pscustomobject]@{Name='Nozomi.la'; Key='Nozomi'}
)
Write-Host ''
Write-Host '========================================' -ForegroundColor Cyan
Write-Host '       BOORU DOWNLOADER — ПОИСК' -ForegroundColor Cyan
Write-Host '       Тестовый режим: без скачивания' -ForegroundColor Yellow
Write-Host '========================================' -ForegroundColor Cyan
for ($i=0; $i -lt $sites.Count; $i++) { Write-Host ("{0}. {1}" -f ($i+1),$sites[$i].Name) }
Write-Host '0. Выход'
$choice = Read-Host 'Выбери сайт'
if ($choice -notmatch '^\d+$' -or [int]$choice -lt 1 -or [int]$choice -gt $sites.Count) { Write-Host 'Выход.'; exit 0 }
$site = $sites[[int]$choice-1]
$defaultTag = switch ($site.Key) { 'e621' {'miyakuroi'} 'Blep' {'distortionai'} 'Floofgens' {'underscore_abby'} 'Gelbooru' {'tomo_'} 'Rule34' {'stickyai'} 'TBIB' {'oosaka_kanagawa'} 'Nozomi' {'artist:As109'} }
Write-Host "Тестовый тег для $($site.Name): $defaultTag" -ForegroundColor DarkGray
$tag = Read-Host 'Введи тег автора (Enter — использовать тестовый тег)'
if ([string]::IsNullOrWhiteSpace($tag)) { $tag = $defaultTag }
$tag = $tag.Trim()
if ($tag -match '[\r\n]') { Write-Host 'Недопустимый ввод.' -ForegroundColor Red; exit 1 }
Write-Host ''
Write-Log "Поиск: сайт=$($site.Name); тег=$tag" 'Cyan'
try {
    switch ($site.Key) {
        'e621' { $result = Search-E621 $tag $Cfg $headers $timeout $limit }
        'Blep' { $result = Search-Danbooru 'blep.art' 'Blep.art' $tag $headers $timeout $limit }
        'Floofgens' { $result = Search-Danbooru 'floofgens.net' 'Floofgens.net' $tag $headers $timeout $limit }
        'Gelbooru' { $result = Search-GelbooruLike 'gelbooru.com' $tag $Cfg 'Gelbooru' $headers $timeout $limit }
        'Rule34' { $result = Search-GelbooruLike 'api.rule34.xxx' $tag $Cfg 'Rule34' $headers $timeout $limit }
        'TBIB' { $result = Search-GelbooruLike 'tbib.org' $tag $Cfg 'TBIB' $headers $timeout $limit }
        'Nozomi' { $result = Search-Nozomi $tag $headers $timeout }
    }
    $folderName = switch ($site.Key) { 'Rule34' {'Rule34'} default {$site.Key} }
    $outDir = Join-Path (Join-Path $ArchiveRoot $folderName) ($tag -replace '[\\/:*?"<>|]','_')
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    if ($result.Sample.Count -gt 0) {
        $rows = foreach ($p in $result.Sample) {
            $id = if ($p.id) {$p.id} elseif ($p.post_id) {$p.post_id} else {''}
            # e621: актуальная структура находится в files.meta / files.original.
            # Старые поля оставлены как fallback для Danbooru-подобных сайтов.
            $url = if ($p.files.original.url) {$p.files.original.url} elseif ($p.file.url) {$p.file.url} elseif ($p.file_url) {$p.file_url} elseif ($p.image) {$p.image} else {''}
            $size = if ($p.files.meta.size) {$p.files.meta.size} elseif ($p.file.size) {$p.file.size} elseif ($p.file_size) {$p.file_size} else {''}
            $ext = if ($p.files.meta.ext) {$p.files.meta.ext} elseif ($p.file.ext) {$p.file.ext} elseif ($p.image) {[IO.Path]::GetExtension([string]$p.image)} else {''}
            $tagText = ''
            if ($p.tags) {
                if ($p.tags -is [string]) { $tagText = $p.tags }
                elseif ($p.tags.PSObject.Properties) {
                    $tagParts = foreach ($prop in $p.tags.PSObject.Properties) { @($prop.Value) }
                    $tagText = (@($tagParts | Where-Object { $_ }) -join ' ')
                } else { $tagText = (@($p.tags) -join ' ') }
            }
            [pscustomobject]@{Id=$id; Ext=$ext; Size=$size; Url=$url; Tags=$tagText}
        }
        $rows | Export-Csv -LiteralPath (Join-Path $outDir '_search_sample.csv') -NoTypeInformation -Encoding UTF8
    }
    $result | Select-Object Count,Method,SearchUrl,Note | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $outDir '_search_result.json') -Encoding UTF8
    Write-Host ''
    Write-Host '------------- РЕЗУЛЬТАТ -------------' -ForegroundColor Cyan
    Write-Host "Сайт:             $($site.Name)"
    Write-Host "Тег:              $tag"
    Write-Host "Найдено:          $($result.Count)"
    Write-Host "Метод:            $($result.Method)"
    if ($result.Note) { Write-Host "Примечание:       $($result.Note)" -ForegroundColor Yellow }
    Write-Host "Примеров в ответе: $(@($result.Sample).Count)"
    Write-Host "Страница сайта:   $($result.SearchUrl)"
    Write-Host "Отчёт:            $outDir"
    Write-Host '-------------------------------------' -ForegroundColor Cyan
    Write-Log "Успех: $($site.Name), тег=$tag, найдено=$($result.Count)" 'Green'
} catch {
    $msg = $_.Exception.Message
    Write-Host ''
    Write-Host "[ОШИБКА ПОИСКА] $msg" -ForegroundColor Red
    Write-Host 'Подробности записаны в Logs\search.log' -ForegroundColor Yellow
    Write-Log "Ошибка: $($site.Name), тег=${tag}: $msg" 'Red'
    if ($site.Key -eq 'Gelbooru') { Write-Host 'Для Gelbooru проверь UserId и ApiKey в config.ini.' -ForegroundColor Yellow }
    if ($site.Key -eq 'TBIB') { Write-Host 'Адаптер TBIB пока экспериментальный: пришли текст ошибки, подстрою запрос под фактический ответ сайта.' -ForegroundColor Yellow }
    if ($site.Key -eq 'Nozomi') { Write-Host 'Nozomi использует бинарные индексы; если HEAD запрещён, пришли текст ошибки — сделаем Range-запрос.' -ForegroundColor Yellow }
}
Write-Host ''
Write-Host 'Это тест поиска. Файлы изображений не скачиваются.' -ForegroundColor Yellow
