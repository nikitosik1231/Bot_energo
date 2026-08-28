[CmdletBinding()]
param(
    [string]$ConfigPath
)

$ErrorActionPreference = 'Stop'
$ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($ScriptDirectory)) { $ScriptDirectory = $PWD.Path }
if ([string]::IsNullOrWhiteSpace($ConfigPath)) { $ConfigPath = Join-Path $ScriptDirectory 'config.json' }

function Read-JsonFile([string]$Path, $Default) {
    if (-not (Test-Path -LiteralPath $Path)) { return $Default }
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
    return $raw | ConvertFrom-Json
}

function Get-Page([string]$Url) {
    # A normal browser-like user agent is needed because the website may reject
    # generic HTTP clients.
    $headers = @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36'; 'Accept-Language' = 'ru-RU,ru;q=0.9' }
    $lastError = $null
    foreach ($timeoutSeconds in @(45, 90)) {
        try {
            return (Invoke-WebRequest -Uri $Url -Headers $headers -MaximumRedirection 5 -TimeoutSec $timeoutSeconds -UseBasicParsing).Content
        } catch {
            $lastError = $_
            if ($timeoutSeconds -eq 45) { Start-Sleep -Seconds 5 }
        }
    }
    throw $lastError
}

function To-AbsoluteUrl([string]$Href, [uri]$BaseUri) {
    if ([string]::IsNullOrWhiteSpace($Href)) { return $null }
    try { return ([uri]::new($BaseUri, $Href)).AbsoluteUri } catch { return $null }
}

function Get-DetailLines([string]$BodyHtml) {
    $withBreaks = [regex]::Replace($BodyHtml, '<br\s*/?>', "`n", 'IgnoreCase')
    $withBreaks = [regex]::Replace($withBreaks, '</p\s*>', "`n", 'IgnoreCase')
    $plain = [System.Net.WebUtility]::HtmlDecode(($withBreaks -replace '<[^>]+>', ''))
    $result = @()
    foreach ($raw in ($plain -split "`r`n|`n|`r")) {
        $line = ($raw -replace '\s+', ' ').Trim()
        if (-not $line) { continue }
        $isArea = $line -match '^(?:г\.\s*|город\s+|[А-ЯЁ][\p{L}\- ]*район(?:\s|,|$)|[А-ЯЁ][\p{L}\- ]*округ(?:\s|,|$))'
        $result += [pscustomobject]@{ Text = $line.Trim(' ', ',', ':', ';'); IsArea = $isArea }
    }
    return $result
}

function Get-ArticleLinks([string]$Html, [uri]$BaseUri) {
    $found = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($m in [regex]::Matches($Html, '<a\b[^>]*href\s*=\s*["''](?<href>[^"'']+)["''][^>]*>(?<text>.*?)</a>', 'IgnoreCase,Singleline')) {
        $title = [System.Net.WebUtility]::HtmlDecode(($m.Groups['text'].Value -replace '<[^>]+>', ' ' -replace '\s+', ' ').Trim())
        $url = To-AbsoluteUrl $m.Groups['href'].Value $BaseUri
        if ($url -and $title) {
            $uri = [uri]$url
            if ($uri.AbsolutePath -match '^/about/news/[^/]+$') { [void]$found.Add($url) }
        }
    }
    return $found
}

function Get-Article([string]$Url) {
    $html = Get-Page $Url
    $titleMatch = [regex]::Match($html, '<h2\b[^>]*itemprop\s*=\s*["'']headline["''][^>]*>(?<value>.*?)</h2>', 'IgnoreCase,Singleline')
    $title = if ($titleMatch.Success) { [System.Net.WebUtility]::HtmlDecode(($titleMatch.Groups['value'].Value -replace '<[^>]+>', ' ' -replace '\s+', ' ').Trim()) } else { $Url }
    $bodyMatch = [regex]::Match($html, '<div\b[^>]*itemprop\s*=\s*["'']articleBody["''][^>]*>(?<value>.*?)</div>', 'IgnoreCase,Singleline')
    $bodyHtml = if ($bodyMatch.Success) { $bodyMatch.Groups['value'].Value } else { '' }
    $text = if ($bodyMatch.Success) { [System.Net.WebUtility]::HtmlDecode(($bodyHtml -replace '<[^>]+>', ' ' -replace '\s+', ' ').Trim()) } else { '' }
    $dateMatch = [regex]::Match($text, '\b\d{1,2}\s+(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+\d{4}\b', 'IgnoreCase')
    $date = if ($dateMatch.Success) { $dateMatch.Value } else { '' }
    $times = @([regex]::Matches($text, '(?:(?:с|до)\s*)?\d{1,2}:\d{2}', 'IgnoreCase') | ForEach-Object { $_.Value.Trim() } | Select-Object -Unique)
    $time = $times -join ', '
    $areaMatch = [regex]::Match($text, 'будет\s+(?:отсутствовать|ограничено)[^.]*?(?:\sв\s|\sпо\s)(?<value>[^.]+)', 'IgnoreCase')
    $area = if ($areaMatch.Success) { $areaMatch.Groups['value'].Value.Trim() } else { '' }
    $areas = @([regex]::Matches($bodyHtml, '<(?:strong|b)\b[^>]*>(?<value>.*?)</(?:strong|b)>', 'IgnoreCase,Singleline') |
        ForEach-Object { [System.Net.WebUtility]::HtmlDecode(($_.Groups['value'].Value -replace '<[^>]+>', ' ' -replace '\s+', ' ').Trim(' ', ',', ':', ';')) } |
        Where-Object { $_ -match '(^г\.\s*|^город\s+|район$)' } |
        Select-Object -Unique)
    $detailLines = Get-DetailLines $bodyHtml
    $attachments = @()
    foreach ($m in [regex]::Matches($html, '(?:href|src)\s*=\s*["''](?<value>[^"'']+\.(?:pdf|xlsx?|docx?|png|jpe?g)(?:\?[^"'']*)?)["'']', 'IgnoreCase')) {
        $attachment = To-AbsoluteUrl $m.Groups['value'].Value ([uri]$Url)
        if ($attachment -and $attachments -notcontains $attachment) { $attachments += $attachment }
    }
    return [pscustomobject]@{ Url = $Url; Title = $title; Text = $text; Date = $date; Time = $time; Area = $area; Areas = $areas; DetailLines = $detailLines; Attachments = $attachments }
}

function Invoke-Telegram([string]$Method, $Payload, $Config) {
    $json = $Payload | ConvertTo-Json -Compress -Depth 6
    return Invoke-RestMethod -Method Post -Uri "https://api.telegram.org/bot$($Config.telegram_bot_token)/$Method" -ContentType 'application/json; charset=utf-8' -Body $json
}

function Split-TelegramLine([string]$Line) {
    if ($Line.Length -le 3600) { return @($Line) }
    $parts = @()
    $remaining = $Line
    while ($remaining.Length -gt 3600) {
        $cut = $remaining.LastIndexOf(' ', 3600)
        if ($cut -lt 1) { $cut = 3600 }
        $parts += $remaining.Substring(0, $cut)
        $remaining = $remaining.Substring($cut).TrimStart()
    }
    $parts += $remaining
    return $parts
}

function Send-TelegramLines([string[]]$Lines, $Config, [long[]]$ChatIds) {
    $messages = @()
    $current = ''
    foreach ($sourceLine in $Lines) {
        foreach ($line in (Split-TelegramLine $sourceLine)) {
            $candidate = if ($current) { $current + "`n" + $line } else { $line }
            if ($candidate.Length -gt 3800 -and $current) {
                $messages += $current
                $current = '📄 <b>Продолжение</b>' + "`n" + $line
            } else {
                $current = $candidate
            }
        }
    }
    if ($current) { $messages += $current }
    foreach ($message in $messages) {
        foreach ($chatId in $ChatIds) {
            Invoke-Telegram 'sendMessage' @{ chat_id = $chatId; text = $message; parse_mode = 'HTML'; disable_web_page_preview = $false } $Config | Out-Null
        }
    }
}

function Send-Telegram([pscustomobject]$Article, $Config, [long[]]$ChatIds) {
    $lines = @('⚡ <b>Крымэнерго — отключения</b>', '', '<b>📌 {0}</b>' -f [System.Net.WebUtility]::HtmlEncode($Article.Title))
    if ($Article.Date) { $lines += '📅 <b>Дата:</b> {0}' -f [System.Net.WebUtility]::HtmlEncode($Article.Date) }
    if ($Article.Time) { $lines += '⏰ <b>Время:</b> {0}' -f [System.Net.WebUtility]::HtmlEncode($Article.Time) }
    $lines += @('', '📝 <b>Детали по территориям:</b>')
    foreach ($detail in $Article.DetailLines) {
        $encoded = [System.Net.WebUtility]::HtmlEncode($detail.Text)
        if ($detail.IsArea) { $lines += '📍 <b>' + $encoded + '</b>' } else { $lines += '▫️ ' + $encoded }
    }
    $lines += @('', '🔗 ' + [System.Net.WebUtility]::HtmlEncode($Article.Url))
    Send-TelegramLines $lines $Config $ChatIds
}

function Receive-Subscribers($Config, $State) {
    if ($null -eq $State.telegram_update_offset) { $State | Add-Member -NotePropertyName telegram_update_offset -NotePropertyValue 0 }
    if ($null -eq $State.subscribers) { $State | Add-Member -NotePropertyName subscribers -NotePropertyValue @() }
    $result = Invoke-Telegram 'getUpdates' @{ offset = $State.telegram_update_offset; timeout = 0; allowed_updates = @('message', 'channel_post', 'my_chat_member') } $Config
    foreach ($update in $result.result) {
        $State.telegram_update_offset = [long]$update.update_id + 1
        $message = $update.message
        $channelPost = $update.channel_post
        $membership = $update.my_chat_member
        if ($channelPost) {
            $chatId = [long]$channelPost.chat.id
            if ($State.subscribers -notcontains $chatId) {
                $isFirstSubscriber = $State.subscribers.Count -eq 0
                $State.subscribers += $chatId
                if ($isFirstSubscriber) { $State.sent_urls = @(); $State.initialized = $false }
            }
            continue
        }
        if ($membership -and $membership.chat.type -eq 'channel' -and $membership.new_chat_member.status -in @('administrator', 'member')) {
            $chatId = [long]$membership.chat.id
            if ($State.subscribers -notcontains $chatId) {
                $isFirstSubscriber = $State.subscribers.Count -eq 0
                $State.subscribers += $chatId
                if ($isFirstSubscriber) { $State.sent_urls = @(); $State.initialized = $false }
            }
            continue
        }
        if ($message -and $message.text -match '^/start(?:\s|$)') {
            $chatId = [long]$message.chat.id
            if ($State.subscribers -notcontains $chatId) {
                $isFirstSubscriber = $State.subscribers.Count -eq 0
                $State.subscribers += $chatId
                if ($isFirstSubscriber) { $State.sent_urls = @(); $State.initialized = $false }
            }
            Invoke-Telegram 'sendMessage' @{ chat_id = $chatId; text = 'Готово. Буду присылать новые новости Крымэнерго с графиками и сообщениями об отключениях.' } $Config | Out-Null
        }
    }
}

function Save-State($State, [string]$Path) {
    $State | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8
}

$config = Read-JsonFile $ConfigPath $null
if (-not $config -or -not $config.telegram_bot_token) {
    throw "Заполните telegram_bot_token в config.json (см. config.example.json)."
}

$statePath = Join-Path $ScriptDirectory 'state.json'
$state = Read-JsonFile $statePath ([pscustomobject]@{ sent_urls = @(); initialized = $false; subscribers = @(); telegram_update_offset = 0 })
if ($null -eq $state.sent_urls) { $state | Add-Member -NotePropertyName sent_urls -NotePropertyValue @() }
Receive-Subscribers $config $state
Save-State $state $statePath
Write-Host "Telegram: subscribers = $($state.subscribers.Count)"

$newsUrl = if ($config.news_url) { $config.news_url } else { 'https://crimea-energy.ru/about/news' }
Write-Host "Reading news: $newsUrl"
$newsHtml = Get-Page $newsUrl
$links = Get-ArticleLinks $newsHtml ([uri]$newsUrl)
Write-Host "News links found: $($links.Count)"
$keywords = @('график', 'отключен', 'обесточ', 'планов', 'электроснабжен', 'энергоснабжен', 'аварийн', 'восстановительн')
$newUrls = @($links | Where-Object { $state.sent_urls -notcontains $_ } | Select-Object -First 3)
[array]::Reverse($newUrls)

foreach ($url in $newUrls) {
    try {
        $article = Get-Article $url
        $haystack = "$($article.Title) $($article.Text)".ToLowerInvariant()
        $isRelevant = $false
        foreach ($keyword in $keywords) { if ($haystack.Contains($keyword)) { $isRelevant = $true; break } }
        if ($isRelevant -and $state.subscribers.Count -gt 0 -and ($state.initialized -or [bool]$config.send_existing_on_first_run)) {
            Send-Telegram $article $config $state.subscribers
            Write-Host "Sent: $($article.Title)"
        }
        $state.sent_urls += $url
    } catch {
        Write-Warning "Не удалось обработать $url : $($_.Exception.Message)"
    }
}

$state.initialized = $true
$state.sent_urls = @($state.sent_urls | Select-Object -Unique | Select-Object -Last 500)
Save-State $state $statePath
Write-Host 'Done.'
