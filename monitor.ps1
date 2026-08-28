[CmdletBinding()]
param(
    [string]$ConfigPath
)

$ErrorActionPreference = 'Stop'

$ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($ScriptDirectory)) {
    $ScriptDirectory = (Get-Location).Path
}

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $ScriptDirectory 'config.json'
}

$DefaultNewsUrl = 'https://crimea-energy.ru/about/news'
$RetryDelaySeconds = 15

$RelevantKeywords = @(
    'график',
    'отключен',
    'отключение',
    'обесточ',
    'планов',
    'электроснабжен',
    'энергоснабжен',
    'аварийн',
    'восстановительн',
    'ремонтн'
)

function Read-JsonFile {
    param([string]$Path, $Default)
    if (-not (Test-Path -LiteralPath $Path)) { return $Default }
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
    return ($raw | ConvertFrom-Json)
}

function Save-JsonFile {
    param($Object, [string]$Path)
    $Object | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Ensure-StateProperties {
    param($State)

    if ($null -eq $State.known_urls) {
        $State | Add-Member -NotePropertyName known_urls -NotePropertyValue @()
    }
    if ($null -eq $State.sent_urls) {
        $State | Add-Member -NotePropertyName sent_urls -NotePropertyValue @()
    }
    if ($null -eq $State.subscribers) {
        $State | Add-Member -NotePropertyName subscribers -NotePropertyValue @()
    }
    if ($null -eq $State.initialized) {
        $State | Add-Member -NotePropertyName initialized -NotePropertyValue $false
    }
    if ($null -eq $State.telegram_update_offset) {
        $State | Add-Member -NotePropertyName telegram_update_offset -NotePropertyValue 0
    }

    return $State
}

function Get-Page {
    param([Parameter(Mandatory = $true)][string]$Url)

    $headers = @{
        'User-Agent'      = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36'
        'Accept'          = 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
        'Accept-Language' = 'ru-RU,ru;q=0.9'
        'Cache-Control'   = 'no-cache'
        'Pragma'          = 'no-cache'
    }

    Write-Host "HTTP GET: $Url"

    $response = Invoke-WebRequest `
        -Uri $Url `
        -Headers $headers `
        -MaximumRedirection 10 `
        -TimeoutSec 90 `
        -UseBasicParsing `
        -ErrorAction Stop

    if ([string]::IsNullOrWhiteSpace($response.Content)) {
        throw 'Сайт вернул пустой ответ.'
    }

    Write-Host "HTTP status: $($response.StatusCode)"
    Write-Host "Downloaded bytes: $($response.RawContentLength)"

    return $response.Content
}

function Get-PageUntilSuccess {
    param([Parameter(Mandatory = $true)][string]$Url)

    $attempt = 0

    while ($true) {
        $attempt++

        try {
            Write-Host "Attempt #$attempt: $Url"
            return Get-Page $Url
        }
        catch {
            Write-Warning "Attempt #$attempt failed: $($_.Exception.Message)"
            Write-Host "Повтор через $RetryDelaySeconds сек."
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

function To-AbsoluteUrl {
    param([string]$Href, [uri]$BaseUri)

    if ([string]::IsNullOrWhiteSpace($Href)) { return $null }

    $Href = [System.Net.WebUtility]::HtmlDecode($Href).Trim()

    try {
        $uri = [uri]::new($BaseUri, $Href)
        if ($uri.Scheme -notin @('http', 'https')) { return $null }
        return $uri.AbsoluteUri
    }
    catch {
        return $null
    }
}

function Convert-HtmlToText {
    param([string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) { return '' }

    $text = $Html
    $text = [regex]::Replace($text, '<br\s*/?>', "`n", 'IgnoreCase')
    $text = [regex]::Replace($text, '</p\s*>', "`n", 'IgnoreCase')
    $text = [regex]::Replace($text, '</div\s*>', "`n", 'IgnoreCase')
    $text = [regex]::Replace($text, '</li\s*>', "`n", 'IgnoreCase')
    $text = [regex]::Replace($text, '<script\b[^>]*>.*?</script>', '', 'IgnoreCase,Singleline')
    $text = [regex]::Replace($text, '<style\b[^>]*>.*?</style>', '', 'IgnoreCase,Singleline')
    $text = $text -replace '<[^>]+>', ' '
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    $text = [regex]::Replace($text, '[ \t]+', ' ')
    $text = [regex]::Replace($text, ' *\n *', "`n")
    return $text.Trim()
}

function Test-ArticleUrl {
    param([string]$Url)

    try {
        $uri = [uri]$Url
        return $uri.AbsolutePath -match '^/about/news/[^/]+$'
    }
    catch {
        return $false
    }
}

function Get-ArticleLinks {
    param([string]$Html, [uri]$BaseUri)

    $result = @()
    $seen = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    $pattern = '<a\b[^>]*href\s*=\s*["''](?<href>[^"'']+)["''][^>]*>(?<text>.*?)</a>'
    $matches = [regex]::Matches($Html, $pattern, 'IgnoreCase,Singleline')

    foreach ($match in $matches) {
        $url = To-AbsoluteUrl $match.Groups['href'].Value $BaseUri
        if (-not $url) { continue }
        if (-not (Test-ArticleUrl $url)) { continue }

        try {
            $uri = [uri]$url
            $url = $uri.GetLeftPart([System.UriPartial]::Path)
        }
        catch { }

        if ($seen.Add($url)) {
            $result += [pscustomobject]@{
                Url = $url
                Title = Convert-HtmlToText $match.Groups['text'].Value
            }
        }
    }

    return @($result)
}

function Get-ArticleTitle {
    param([string]$Html)

    $patterns = @(
        '<h1[^>]*>(?<value>.*?)</h1>',
        '<h2[^>]*itemprop\s*=\s*["'']headline["''][^>]*>(?<value>.*?)</h2>',
        '<title[^>]*>(?<value>.*?)</title>'
    )

    foreach ($pattern in $patterns) {
        $match = [regex]::Match($Html, $pattern, 'IgnoreCase,Singleline')
        if ($match.Success) {
            $title = Convert-HtmlToText $match.Groups['value'].Value
            if ($title -and $title -ne 'Новости') {
                return $title
            }
        }
    }

    return 'Новости'
}

function Get-ArticleBody {
    param([string]$Html)

    $patterns = @(
        '<div[^>]*itemprop\s*=\s*["'']articleBody["''][^>]*>(?<value>.*?)</div>',
        '<article[^>]*>(?<value>.*?)</article>'
    )

    foreach ($pattern in $patterns) {
        $match = [regex]::Match($Html, $pattern, 'IgnoreCase,Singleline')
        if ($match.Success -and $match.Groups['value'].Value) {
            return $match.Groups['value'].Value
        }
    }

    $bodyMatch = [regex]::Match($Html, '<body[^>]*>(?<value>.*?)</body>', 'IgnoreCase,Singleline')
    if ($bodyMatch.Success) { return $bodyMatch.Groups['value'].Value }

    return $Html
}

function Get-DetailLines {
    param([string]$BodyHtml)

    $plain = $BodyHtml
    $plain = [regex]::Replace($plain, '<br\s*/?>', "`n", 'IgnoreCase')
    $plain = [regex]::Replace($plain, '</p\s*>', "`n", 'IgnoreCase')
    $plain = [regex]::Replace($plain, '</div\s*>', "`n", 'IgnoreCase')
    $plain = [regex]::Replace($plain, '</li\s*>', "`n", 'IgnoreCase')
    $plain = Convert-HtmlToText $plain

    $result = @()

    foreach ($raw in ($plain -split "`r`n|`n|`r")) {
        $line = ($raw -replace '\s+', ' ').Trim()
        if (-not $line) { continue }

        $isArea =
            $line -match '^г\.\s*' -or
            $line -match '^город\s+' -or
            $line -match 'район(?:а)?$' -or
            $line -match 'округ(?:а)?$' -or
            $line -match '^с\.\s*' -or
            $line -match '^пгт\.\s*'

        $result += [pscustomobject]@{
            Text   = $line.Trim(' ', ',', ':', ';')
            IsArea = $isArea
        }
    }

    return @($result)
}

function Get-Article {
    param([string]$Url)

    $html = Get-PageUntilSuccess $Url
    $title = Get-ArticleTitle $html
    $bodyHtml = Get-ArticleBody $html
    $text = Convert-HtmlToText $bodyHtml

    $dateMatch = [regex]::Match(
        $text,
        '\b\d{1,2}\s+(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+\d{4}\b',
        'IgnoreCase'
    )

    $date = if ($dateMatch.Success) { $dateMatch.Value } else { '' }

    $times = @(
        [regex]::Matches($text, '\b\d{1,2}:\d{2}\b') |
        ForEach-Object { $_.Value } |
        Select-Object -Unique
    )

    $time = $times -join '–'

    $areaMatch = [regex]::Match(
        $text,
        '(?:г\.|город|с\.|пгт\.|район|округ)\s*[^,.;\n]{1,100}',
        'IgnoreCase'
    )

    $area = if ($areaMatch.Success) { $areaMatch.Value.Trim() } else { '' }

    return [pscustomobject]@{
        Url         = $Url
        Title       = $title
        Text        = $text
        Date        = $date
        Time        = $time
        Area        = $area
        DetailLines = @(Get-DetailLines $bodyHtml)
    }
}

function Invoke-Telegram {
    param([string]$Method, $Payload, $Config)

    $token = [string]$Config.telegram_bot_token
    if ([string]::IsNullOrWhiteSpace($token)) {
        throw 'telegram_bot_token не указан.'
    }

    $json = $Payload | ConvertTo-Json -Compress -Depth 10
    $uri = "https://api.telegram.org/bot$token/$Method"

    return Invoke-RestMethod `
        -Method Post `
        -Uri $uri `
        -ContentType 'application/json; charset=utf-8' `
        -Body $json `
        -TimeoutSec 60 `
        -ErrorAction Stop
}

function Receive-Subscribers {
    param($Config, $State)

    try {
        $response = Invoke-Telegram `
            'getUpdates' `
            @{
                offset = [long]$State.telegram_update_offset
                timeout = 0
                allowed_updates = @('message','channel_post','my_chat_member')
            } `
            $Config

        foreach ($update in @($response.result)) {
            $State.telegram_update_offset = [long]$update.update_id + 1

            if ($update.message) {
                $chatId = [long]$update.message.chat.id
                $text = [string]$update.message.text

                if ($text -match '^/start(?:\s|$)') {
                    if ($State.subscribers -notcontains $chatId) {
                        $State.subscribers += $chatId
                    }

                    Invoke-Telegram `
                        'sendMessage' `
                        @{
                            chat_id = $chatId
                            text = 'Готово. Буду присылать новые публикации Крымэнерго об отключениях.'
                        } `
                        $Config |
                        Out-Null
                }

                if ($text -match '^/stop(?:\s|$)') {
                    $State.subscribers = @(
                        $State.subscribers |
                        Where-Object { $_ -ne $chatId }
                    )

                    Invoke-Telegram `
                        'sendMessage' `
                        @{
                            chat_id = $chatId
                            text = 'Уведомления отключены.'
                        } `
                        $Config |
                        Out-Null
                }
            }

            if ($update.channel_post) {
                $chatId = [long]$update.channel_post.chat.id
                if ($State.subscribers -notcontains $chatId) {
                    $State.subscribers += $chatId
                }
            }

            if ($update.my_chat_member) {
                $member = $update.my_chat_member

                if (
                    $member.chat.type -in @('channel','group','supergroup') -and
                    $member.new_chat_member.status -in @('member','administrator')
                ) {
                    $chatId = [long]$member.chat.id
                    if ($State.subscribers -notcontains $chatId) {
                        $State.subscribers += $chatId
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Telegram getUpdates: $($_.Exception.Message)"
    }
}

function Split-TelegramLine {
    param([string]$Line)

    if ($Line.Length -le 3600) { return @($Line) }

    $parts = @()
    $remaining = $Line

    while ($remaining.Length -gt 3600) {
        $cut = $remaining.LastIndexOf(' ', 3600)
        if ($cut -lt 1) { $cut = 3600 }

        $parts += $remaining.Substring(0, $cut)
        $remaining = $remaining.Substring($cut).TrimStart()
    }

    if ($remaining) { $parts += $remaining }
    return @($parts)
}

function Send-TelegramLines {
    param(
        [string[]]$Lines,
        $Config,
        [long[]]$ChatIds
    )

    $messages = @()
    $current = ''

    foreach ($sourceLine in $Lines) {
        foreach ($line in (Split-TelegramLine $sourceLine)) {
            $candidate = if ($current) {
                $current + "`n" + $line
            }
            else {
                $line
            }

            if ($candidate.Length -gt 3800 -and $current) {
                $messages += $current
                $current = "📄 <b>Продолжение</b>`n$line"
            }
            else {
                $current = $candidate
            }
        }
    }

    if ($current) { $messages += $current }

    foreach ($message in $messages) {
        foreach ($chatId in $ChatIds) {
            Invoke-Telegram `
                'sendMessage' `
                @{
                    chat_id = $chatId
                    text = $message
                    parse_mode = 'HTML'
                    disable_web_page_preview = $false
                } `
                $Config |
                Out-Null
        }
    }
}

# Старое оформление пользователя.
function Send-Telegram {
    param(
        [pscustomobject]$Article,
        $Config,
        [long[]]$ChatIds
    )

    $lines = @(
        '⚡ <b>Крымэнерго — отключения</b>',
        '',
        ('<b>📌 {0}</b>' -f
            [System.Net.WebUtility]::HtmlEncode($Article.Title))
    )

    if ($Article.Date) {
        $lines += (
            '📅 <b>Дата:</b> {0}' -f
            [System.Net.WebUtility]::HtmlEncode($Article.Date)
        )
    }

    if ($Article.Time) {
        $lines += (
            '⏰ <b>Время:</b> {0}' -f
            [System.Net.WebUtility]::HtmlEncode($Article.Time)
        )
    }

    $lines += @('', '📝 <b>Детали по территориям:</b>')

    foreach ($detail in $Article.DetailLines) {
        $encoded = [System.Net.WebUtility]::HtmlEncode($detail.Text)

        if ($detail.IsArea) {
            $lines += "📍 <b>$encoded</b>"
        }
        else {
            $lines += "▫️ $encoded"
        }
    }

    $lines += @(
        '',
        ('🔗 <a href="{0}">Открыть публикацию</a>' -f
            [System.Net.WebUtility]::HtmlEncode($Article.Url))
    )

    Send-TelegramLines $lines $Config $ChatIds
}

# ============================================================
# MAIN
# ============================================================

Write-Host ''
Write-Host '========================================'
Write-Host ' Crimea Energy Telegram Monitor'
Write-Host '========================================'

$config = Read-JsonFile $ConfigPath $null
if (-not $config) {
    throw "Не найден config.json: $ConfigPath"
}

if ([string]::IsNullOrWhiteSpace([string]$config.telegram_bot_token)) {
    throw 'В config.json отсутствует telegram_bot_token.'
}

$statePath = Join-Path $ScriptDirectory 'state.json'

$state = Read-JsonFile `
    $statePath `
    ([pscustomobject]@{
        known_urls = @()
        sent_urls = @()
        initialized = $false
        subscribers = @()
        telegram_update_offset = 0
    })

$state = Ensure-StateProperties $state

Receive-Subscribers $config $state
Save-JsonFile $state $statePath

Write-Host "Telegram subscribers: $($state.subscribers.Count)"

$newsUrl = [string]$config.news_url
if ([string]::IsNullOrWhiteSpace($newsUrl)) {
    $newsUrl = $DefaultNewsUrl
}

$newsUri = [uri]$newsUrl
Write-Host "News URL: $newsUrl"

# Страница новостей загружается один раз за запуск.
$newsHtml = Get-PageUntilSuccess $newsUrl
$links = @(Get-ArticleLinks $newsHtml $newsUri)

Write-Host "Article links found: $($links.Count)"

if ($links.Count -eq 0) {
    Write-Warning 'На странице не найдено ссылок на статьи.'
    Save-JsonFile $state $statePath
    exit 0
}

$knownSet = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

foreach ($url in @($state.known_urls)) {
    [void]$knownSet.Add([string]$url)
}

$sentSet = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

foreach ($url in @($state.sent_urls)) {
    [void]$sentSet.Add([string]$url)
}

# ============================================================
# FIRST RUN:
# только 3 последние релевантные статьи.
# ============================================================

if (-not [bool]$state.initialized) {

    Write-Host ''
    Write-Host 'FIRST RUN: ищем 3 последние релевантные публикации.'

    $selected = @()

    # Сайт отдаёт новости от новых к старым.
    # Ищем три последние, открывая статьи последовательно.
    foreach ($link in $links) {

        if ($selected.Count -ge 3) { break }

        Write-Host ''
        Write-Host "Checking candidate: $($link.Url)"

        # Если статья не открывается, Get-PageUntilSuccess
        # продолжает попытки до успешного получения.
        $article = Get-Article $link.Url

        $haystack = "$($article.Title) $($article.Text)".ToLowerInvariant()
        $isRelevant = $false

        foreach ($keyword in $RelevantKeywords) {
            if ($haystack.Contains($keyword.ToLowerInvariant())) {
                $isRelevant = $true
                break
            }
        }

        Write-Host "Title: $($article.Title)"
        Write-Host "Relevant: $isRelevant"

        [void]$knownSet.Add($link.Url)

        if ($isRelevant) {
            $selected += $article
        }
    }

    # Выбраны от новой к старой — отправляем от старой к новой.
    [array]::Reverse($selected)

    Write-Host ''
    Write-Host "FIRST RUN selected: $($selected.Count)"

    foreach ($article in $selected) {

        if ($sentSet.Contains($article.Url)) { continue }

        if ($state.subscribers.Count -gt 0) {
            Send-Telegram `
                $article `
                $config `
                ([long[]]$state.subscribers)

            Write-Host "Telegram: SENT - $($article.Title)"
        }
        else {
            Write-Warning 'Нет подписчиков Telegram. Новость не отправлена.'
        }

        [void]$sentSet.Add($article.Url)
        [void]$knownSet.Add($article.Url)

        $state.known_urls = @($knownSet | Select-Object -Last 1000)
        $state.sent_urls = @($sentSet | Select-Object -Last 500)

        # Сохраняем после каждой статьи.
        Save-JsonFile $state $statePath
    }

    # Все публикации, существовавшие на момент первого запуска,
    # считаются известными, но отправляются только выбранные три.
    foreach ($link in $links) {
        [void]$knownSet.Add($link.Url)
    }

    $state.known_urls = @($knownSet | Select-Object -Last 1000)
    $state.sent_urls = @($sentSet | Select-Object -Last 500)
    $state.initialized = $true

    Save-JsonFile $state $statePath

    Write-Host 'FIRST RUN COMPLETE.'
    exit 0
}

# ============================================================
# NORMAL RUN:
# только URL, которых раньше не было.
# ============================================================

Write-Host ''
Write-Host 'NORMAL RUN: ищем только новые публикации.'

$newLinks = @(
    $links |
    Where-Object { -not $knownSet.Contains($_.Url) }
)

Write-Host "New URLs: $($newLinks.Count)"

if ($newLinks.Count -eq 0) {
    foreach ($link in $links) {
        [void]$knownSet.Add($link.Url)
    }

    $state.known_urls = @($knownSet | Select-Object -Last 1000)
    Save-JsonFile $state $statePath

    Write-Host 'No new publications.'
    exit 0
}

# На сайте новые идут первыми. Публикуем старые новые -> новые новые.
[array]::Reverse($newLinks)

foreach ($link in $newLinks) {

    Write-Host ''
    Write-Host '----------------------------------------'
    Write-Host "NEW ARTICLE: $($link.Url)"

    # Не переходим к следующей статье, пока эту не получили.
    $article = Get-Article $link.Url

    $haystack = "$($article.Title) $($article.Text)".ToLowerInvariant()
    $isRelevant = $false

    foreach ($keyword in $RelevantKeywords) {
        if ($haystack.Contains($keyword.ToLowerInvariant())) {
            $isRelevant = $true
            break
        }
    }

    Write-Host "Title: $($article.Title)"
    Write-Host "Relevant: $isRelevant"

    if ($isRelevant) {
        if ($state.subscribers.Count -gt 0) {
            Send-Telegram `
                $article `
                $config `
                ([long[]]$state.subscribers)

            Write-Host 'Telegram: SENT'
        }
        else {
            Write-Warning 'Нет подписчиков Telegram. Новость не отправлена.'
        }

        [void]$sentSet.Add($link.Url)
    }
    else {
        Write-Host 'Telegram: NOT SENT (not relevant)'
    }

    [void]$knownSet.Add($link.Url)

    $state.known_urls = @($knownSet | Select-Object -Last 1000)
    $state.sent_urls = @($sentSet | Select-Object -Last 500)

    # Сохраняем после каждой обработанной публикации.
    Save-JsonFile $state $statePath
}

Write-Host ''
Write-Host 'NORMAL RUN COMPLETE.'
