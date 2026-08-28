[CmdletBinding()]
param(
    [string]$ConfigPath
)

$ErrorActionPreference = 'Stop'

# ============================================================
# КРЫМЭНЕРГО MONITOR
# ============================================================

$ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

if ([string]::IsNullOrWhiteSpace($ScriptDirectory)) {
    $ScriptDirectory = (Get-Location).Path
}

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $ScriptDirectory 'config.json'
}

# ------------------------------------------------------------
# НАСТРОЙКИ
# ------------------------------------------------------------

$DefaultNewsUrl = 'https://crimea-energy.ru/about/news'

$AllowedArticlePaths = @(
    '/about/news/',
    '/consumers/cserv/classifieds/'
)

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

# ------------------------------------------------------------
# JSON
# ------------------------------------------------------------

function Read-JsonFile {
    param(
        [string]$Path,
        $Default
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $Default
    }

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8

    if ([string]::IsNullOrWhiteSpace($raw)) {
        return $Default
    }

    return ($raw | ConvertFrom-Json)
}

function Save-JsonFile {
    param(
        $Object,
        [string]$Path
    )

    $Object |
        ConvertTo-Json -Depth 10 |
        Set-Content -LiteralPath $Path -Encoding UTF8
}

# ------------------------------------------------------------
# HTTP
# ------------------------------------------------------------

function Get-Page {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    $userAgents = @(
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36',
        'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36'
    )

    $lastError = $null

    foreach ($userAgent in $userAgents) {

        try {

            Write-Host "HTTP GET: $Url"
            Write-Host "User-Agent: $userAgent"

            $headers = @{
                'User-Agent'      = $userAgent
                'Accept'          = 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
                'Accept-Language' = 'ru-RU,ru;q=0.9,en;q=0.5'
                'Cache-Control'   = 'no-cache'
                'Pragma'          = 'no-cache'
            }

            $response = Invoke-WebRequest `
                -Uri $Url `
                -Headers $headers `
                -MaximumRedirection 10 `
                -TimeoutSec 60 `
                -UseBasicParsing

            Write-Host "HTTP status: $($response.StatusCode)"
            Write-Host "Downloaded bytes: $($response.RawContentLength)"

            if ([string]::IsNullOrWhiteSpace($response.Content)) {
                throw "Сайт вернул пустой ответ."
            }

            return $response.Content

        }
        catch {

            $lastError = $_

            Write-Warning "Ошибка HTTP: $($_.Exception.Message)"

            Start-Sleep -Seconds 3
        }
    }

    throw $lastError
}

# ------------------------------------------------------------
# URL
# ------------------------------------------------------------

function To-AbsoluteUrl {
    param(
        [string]$Href,
        [uri]$BaseUri
    )

    if ([string]::IsNullOrWhiteSpace($Href)) {
        return $null
    }

    $Href = [System.Net.WebUtility]::HtmlDecode($Href).Trim()

    try {
        $uri = [uri]::new($BaseUri, $Href)

        if ($uri.Scheme -notin @('http', 'https')) {
            return $null
        }

        return $uri.AbsoluteUri
    }
    catch {
        return $null
    }
}

# ------------------------------------------------------------
# HTML → TEXT
# ------------------------------------------------------------

function Convert-HtmlToText {
    param(
        [string]$Html
    )

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ''
    }

    $text = $Html

    # Переносы
    $text = [regex]::Replace(
        $text,
        '<br\s*/?>',
        "`n",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    $text = [regex]::Replace(
        $text,
        '</p\s*>',
        "`n",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    $text = [regex]::Replace(
        $text,
        '</div\s*>',
        "`n",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    $text = [regex]::Replace(
        $text,
        '</li\s*>',
        "`n",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    # Убираем скрипты
    $text = [regex]::Replace(
        $text,
        '<script\b[^>]*>.*?</script>',
        '',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    # Убираем CSS
    $text = [regex]::Replace(
        $text,
        '<style\b[^>]*>.*?</style>',
        '',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    # HTML
    $text = $text -replace '<[^>]+>', ' '

    # HTML entities
    $text = [System.Net.WebUtility]::HtmlDecode($text)

    # Нормализация пробелов
    $text = [regex]::Replace($text, '[ \t]+', ' ')
    $text = [regex]::Replace($text, '\n[ \t]+', "`n")

    return $text.Trim()
}

# ------------------------------------------------------------
# ПРОВЕРКА ССЫЛКИ
# ------------------------------------------------------------

function Test-ArticleUrl {
    param(
        [string]$Url
    )

    try {

        $uri = [uri]$Url

        foreach ($allowedPath in $AllowedArticlePaths) {

            if ($uri.AbsolutePath.StartsWith(
                $allowedPath,
                [System.StringComparison]::OrdinalIgnoreCase
            )) {
                return $true
            }
        }

        return $false
    }
    catch {
        return $false
    }
}

# ------------------------------------------------------------
# ПОИСК ССЫЛОК НА НОВОСТИ
# ------------------------------------------------------------

function Get-ArticleLinks {
    param(
        [string]$Html,
        [uri]$BaseUri
    )

    $found = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    $pattern = '<a\b[^>]*href\s*=\s*["''](?<href>[^"'']+)["''][^>]*>(?<text>.*?)</a>'

    $matches = [regex]::Matches(
        $Html,
        $pattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    foreach ($match in $matches) {

        $href = $match.Groups['href'].Value
        $title = Convert-HtmlToText $match.Groups['text'].Value

        $url = To-AbsoluteUrl $href $BaseUri

        if (-not $url) {
            continue
        }

        if (-not $title) {
            continue
        }

        if (-not (Test-ArticleUrl $url)) {
            continue
        }

        [void]$found.Add($url)
    }

    return @($found)
}

# ------------------------------------------------------------
# RSS
# ------------------------------------------------------------

function Get-RssArticles {
    param(
        [string]$Url
    )

    $html = Get-Page $Url

    if ([string]::IsNullOrWhiteSpace($html)) {
        return @()
    }

    try {

        [xml]$xml = $html

        $result = @()

        # RSS
        if ($xml.rss.channel.item) {

            foreach ($item in @($xml.rss.channel.item)) {

                $link = [string]$item.link

                if (-not $link) {
                    continue
                }

                if (-not (Test-ArticleUrl $link)) {
                    continue
                }

                $result += [pscustomobject]@{
                    Url      = $link
                    Title    = Convert-HtmlToText ([string]$item.title)
                    BodyHtml = [string]$item.description
                }
            }
        }

        # Atom
        if ($xml.feed.entry) {

            foreach ($entry in @($xml.feed.entry)) {

                $link = ''

                if ($entry.link.href) {
                    $link = [string]$entry.link.href
                }

                if (-not $link) {
                    continue
                }

                if (-not (Test-ArticleUrl $link)) {
                    continue
                }

                $result += [pscustomobject]@{
                    Url      = $link
                    Title    = Convert-HtmlToText ([string]$entry.title)
                    BodyHtml = [string]$entry.content
                }
            }
        }

        return @(
            $result |
            Sort-Object Url -Unique
        )
    }
    catch {

        throw "RSS/XML не удалось разобрать: $($_.Exception.Message)"
    }
}

# ------------------------------------------------------------
# ЗАГОЛОВОК СТРАНИЦЫ
# ------------------------------------------------------------

function Get-ArticleTitle {
    param(
        [string]$Html
    )

    $patterns = @(
        '<h1[^>]*>(?<value>.*?)</h1>',
        '<h2[^>]*itemprop\s*=\s*["'']headline["''][^>]*>(?<value>.*?)</h2>',
        '<title[^>]*>(?<value>.*?)</title>'
    )

    foreach ($pattern in $patterns) {

        $match = [regex]::Match(
            $Html,
            $pattern,
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
            [System.Text.RegularExpressions.RegexOptions]::Singleline
        )

        if ($match.Success) {

            $title = Convert-HtmlToText $match.Groups['value'].Value

            if ($title) {
                return $title
            }
        }
    }

    return ''
}

# ------------------------------------------------------------
# BODY СТАТЬИ
# ------------------------------------------------------------

function Get-ArticleBody {
    param(
        [string]$Html
    )

    $patterns = @(
        '<div[^>]*itemprop\s*=\s*["'']articleBody["''][^>]*>(?<value>.*?)</div>',
        '<div[^>]*class\s*=\s*["''][^"'']*(?:item-page|article|news)[^"'']*["''][^>]*>(?<value>.*?)</div>'
    )

    foreach ($pattern in $patterns) {

        $match = [regex]::Match(
            $Html,
            $pattern,
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
            [System.Text.RegularExpressions.RegexOptions]::Singleline
        )

        if ($match.Success) {

            $body = $match.Groups['value'].Value

            if (-not [string]::IsNullOrWhiteSpace($body)) {
                return $body
            }
        }
    }

    # Если специальный контейнер не найден,
    # берём всё между BODY.
    $bodyMatch = [regex]::Match(
        $Html,
        '<body[^>]*>(?<value>.*?)</body>',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    if ($bodyMatch.Success) {
        return $bodyMatch.Groups['value'].Value
    }

    return $Html
}

# ------------------------------------------------------------
# СТРОКИ АДРЕСОВ
# ------------------------------------------------------------

function Get-DetailLines {
    param(
        [string]$BodyHtml
    )

    $withBreaks = $BodyHtml

    $withBreaks = [regex]::Replace(
        $withBreaks,
        '<br\s*/?>',
        "`n",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    $withBreaks = [regex]::Replace(
        $withBreaks,
        '</p\s*>',
        "`n",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    $withBreaks = [regex]::Replace(
        $withBreaks,
        '</div\s*>',
        "`n",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    $withBreaks = [regex]::Replace(
        $withBreaks,
        '</li\s*>',
        "`n",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    $plain = Convert-HtmlToText $withBreaks

    $result = @()

    foreach ($raw in ($plain -split "`r`n|`n|`r")) {

        $line = ($raw -replace '\s+', ' ').Trim()

        if (-not $line) {
            continue
        }

        # Определяем территорию.
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

# ------------------------------------------------------------
# ИЗВЛЕЧЕНИЕ ВЛОЖЕНИЙ
# ------------------------------------------------------------

function Get-Attachments {
    param(
        [string]$Html,
        [uri]$BaseUri
    )

    $result = @()

    $pattern = '(?:href|src)\s*=\s*["''](?<value>[^"'']+)["'']'

    foreach ($match in [regex]::Matches(
        $Html,
        $pattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )) {

        $url = To-AbsoluteUrl $match.Groups['value'].Value $BaseUri

        if (-not $url) {
            continue
        }

        if ($url -match '\.(pdf|xlsx?|docx?|png|jpe?g)(?:\?.*)?$') {

            if ($result -notcontains $url) {
                $result += $url
            }
        }
    }

    return @($result)
}

# ------------------------------------------------------------
# РАЗБОР СТАТЬИ
# ------------------------------------------------------------

function Get-Article {
    param(
        [string]$Url,
        [string]$RssTitle = '',
        [string]$RssBodyHtml = ''
    )

    $html = ''
    $title = ''
    $bodyHtml = ''

    # Если есть данные RSS — используем их,
    # но всё равно пытаемся получить нормальную страницу.
    if ($RssBodyHtml) {
        $title = $RssTitle
        $bodyHtml = $RssBodyHtml
    }

    try {

        $html = Get-Page $Url

        $pageTitle = Get-ArticleTitle $html

        if ($pageTitle) {
            $title = $pageTitle
        }

        $pageBody = Get-ArticleBody $html

        if ($pageBody) {
            $bodyHtml = $pageBody
        }
    }
    catch {

        if (-not $bodyHtml) {
            throw
        }

        Write-Warning "Не удалось открыть статью $Url. Использую данные RSS."
    }

    if (-not $title) {
        $title = $Url
    }

    $text = Convert-HtmlToText $bodyHtml

    # Дата.
    $dateMatch = [regex]::Match(
        $text,
        '\b\d{1,2}\s+(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+\d{4}\b',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    $date = ''

    if ($dateMatch.Success) {
        $date = $dateMatch.Value
    }

    # Время.
    $times = @(
        [regex]::Matches(
            $text,
            '\b\d{1,2}:\d{2}\b'
        ) |
        ForEach-Object {
            $_.Value
        } |
        Select-Object -Unique
    )

    $time = $times -join '–'

    # Территория.
    $area = ''

    $areaMatch = [regex]::Match(
        $text,
        '(?:г\.|город|с\.|пгт\.|район|округ)\s*[^,.;\n]{1,100}',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    if ($areaMatch.Success) {
        $area = $areaMatch.Value.Trim()
    }

    $detailLines = Get-DetailLines $bodyHtml

    $attachments = Get-Attachments `
        $html `
        ([uri]$Url)

    return [pscustomobject]@{
        Url          = $Url
        Title        = $title
        Text         = $text
        Date         = $date
        Time         = $time
        Area         = $area
        DetailLines  = $detailLines
        Attachments  = $attachments
    }
}

# ------------------------------------------------------------
# TELEGRAM
# ------------------------------------------------------------

function Invoke-Telegram {
    param(
        [string]$Method,
        $Payload,
        $Config
    )

    $token = [string]$Config.telegram_bot_token

    if ([string]::IsNullOrWhiteSpace($token)) {
        throw 'telegram_bot_token не указан.'
    }

    $json = $Payload |
        ConvertTo-Json -Compress -Depth 10

    $uri = "https://api.telegram.org/bot$token/$Method"

    return Invoke-RestMethod `
        -Method Post `
        -Uri $uri `
        -ContentType 'application/json; charset=utf-8' `
        -Body $json
}

# ------------------------------------------------------------
# TELEGRAM — ПОЛУЧЕНИЕ ПОДПИСЧИКОВ
# ------------------------------------------------------------

function Receive-Subscribers {
    param(
        $Config,
        $State
    )

    if ($null -eq $State.telegram_update_offset) {
        $State |
            Add-Member `
                -NotePropertyName telegram_update_offset `
                -NotePropertyValue 0
    }

    if ($null -eq $State.subscribers) {
        $State |
            Add-Member `
                -NotePropertyName subscribers `
                -NotePropertyValue @()
    }

    try {

        $response = Invoke-Telegram `
            'getUpdates' `
            @{
                offset = [long]$State.telegram_update_offset
                timeout = 0
                allowed_updates = @(
                    'message',
                    'channel_post',
                    'my_chat_member'
                )
            } `
            $Config

        if (-not $response.ok) {
            Write-Warning "Telegram getUpdates вернул ошибку."
            return
        }

        foreach ($update in @($response.result)) {

            $State.telegram_update_offset =
                [long]$update.update_id + 1

            # Обычное сообщение.
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
                            text = @"
Готово.

Я буду отслеживать новые публикации Крымэнерго об отключениях.

Для проверки подписки:
 /start

Для остановки уведомлений:
 /stop
"@
                        } `
                        $Config |
                        Out-Null
                }

                if ($text -match '^/stop(?:\s|$)') {

                    $State.subscribers = @(
                        $State.subscribers |
                        Where-Object {
                            $_ -ne $chatId
                        }
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

            # Если бот добавлен в канал.
            if ($update.channel_post) {

                $chatId = [long]$update.channel_post.chat.id

                if ($State.subscribers -notcontains $chatId) {
                    $State.subscribers += $chatId
                }
            }

            # Если бота добавили в канал/группу.
            if ($update.my_chat_member) {

                $member = $update.my_chat_member
                $chat = $member.chat

                if (
                    $chat.type -in @('channel', 'group', 'supergroup') -and
                    $member.new_chat_member.status -in @(
                        'member',
                        'administrator'
                    )
                ) {

                    $chatId = [long]$chat.id

                    if ($State.subscribers -notcontains $chatId) {
                        $State.subscribers += $chatId
                    }
                }
            }
        }
    }
    catch {

        Write-Warning "Ошибка Telegram: $($_.Exception.Message)"
    }
}

# ------------------------------------------------------------
# TELEGRAM — ОТПРАВКА
# ------------------------------------------------------------

function Split-TelegramText {
    param(
        [string]$Text
    )

    $maxLength = 3900

    if ($Text.Length -le $maxLength) {
        return @($Text)
    }

    $parts = @()

    $remaining = $Text

    while ($remaining.Length -gt $maxLength) {

        $cut = $remaining.LastIndexOf(
            "`n",
            $maxLength
        )

        if ($cut -lt 500) {
            $cut = $remaining.LastIndexOf(
                ' ',
                $maxLength
            )
        }

        if ($cut -lt 1) {
            $cut = $maxLength
        }

        $parts += $remaining.Substring(0, $cut)

        $remaining =
            $remaining.Substring($cut).TrimStart()
    }

    if ($remaining) {
        $parts += $remaining
    }

    return @($parts)
}

function Send-TelegramArticle {
    param(
        [pscustomobject]$Article,
        $Config,
        [long[]]$ChatIds
    )

    $title = [System.Net.WebUtility]::HtmlEncode(
        $Article.Title
    )

    $lines = @()

    $lines += '⚡ <b>Крымэнерго — отключения</b>'
    $lines += ''
    $lines += "📌 <b>$title</b>"

    if ($Article.Date) {
        $date =
            [System.Net.WebUtility]::HtmlEncode(
                $Article.Date
            )

        $lines += "📅 <b>Дата:</b> $date"
    }

    if ($Article.Time) {
        $time =
            [System.Net.WebUtility]::HtmlEncode(
                $Article.Time
            )

        $lines += "⏰ <b>Время:</b> $time"
    }

    if ($Article.Area) {

        $area =
            [System.Net.WebUtility]::HtmlEncode(
                $Article.Area
            )

        $lines += "📍 <b>$area</b>"
    }

    $lines += ''
    $lines += '📝 <b>Информация:</b>'

    foreach ($detail in $Article.DetailLines) {

        $text =
            [System.Net.WebUtility]::HtmlEncode(
                $detail.Text
            )

        if ($detail.IsArea) {
            $lines += "📍 <b>$text</b>"
        }
        else {
            $lines += "▫️ $text"
        }
    }

    $lines += ''
    $lines += (
        '🔗 <a href="{0}">Открыть публикацию</a>' -f
        [System.Net.WebUtility]::HtmlEncode($Article.Url)
    )

    $fullText = $lines -join "`n"

    $parts = Split-TelegramText $fullText

    foreach ($part in $parts) {

        foreach ($chatId in $ChatIds) {

            try {

                Invoke-Telegram `
                    'sendMessage' `
                    @{
                        chat_id = $chatId
                        text = $part
                        parse_mode = 'HTML'
                        disable_web_page_preview = $true
                    } `
                    $Config |
                    Out-Null

                Start-Sleep -Milliseconds 300
            }
            catch {

                Write-Warning `
                    "Ошибка отправки Telegram $chatId : $($_.Exception.Message)"
            }
        }
    }
}

# ============================================================
# MAIN
# ============================================================

Write-Host ''
Write-Host '========================================'
Write-Host ' Crimea Energy Telegram Monitor'
Write-Host '========================================'
Write-Host ''

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------

$config = Read-JsonFile `
    $ConfigPath `
    $null

if (-not $config) {

    throw @"
Не найден config.json.

Создай файл:

$ConfigPath

Пример:

{
  "telegram_bot_token": "ТОКЕН_ОТ_BOTFATHER",
  "news_url": "https://crimea-energy.ru/about/news",
  "send_existing_on_first_run": false
}
"@
}

if (
    [string]::IsNullOrWhiteSpace(
        [string]$config.telegram_bot_token
    )
) {
    throw 'В config.json отсутствует telegram_bot_token.'
}

# ------------------------------------------------------------
# STATE
# ------------------------------------------------------------

$statePath = Join-Path `
    $ScriptDirectory `
    'state.json'

$state = Read-JsonFile `
    $statePath `
    ([pscustomobject]@{
        sent_urls = @()
        initialized = $false
        subscribers = @()
        telegram_update_offset = 0
    })

if ($null -eq $state.sent_urls) {
    $state |
        Add-Member `
            -NotePropertyName sent_urls `
            -NotePropertyValue @()
}

if ($null -eq $state.subscribers) {
    $state |
        Add-Member `
            -NotePropertyName subscribers `
            -NotePropertyValue @()
}

# ------------------------------------------------------------
# TELEGRAM
# ------------------------------------------------------------

Receive-Subscribers `
    $config `
    $state

Save-JsonFile `
    $state `
    $statePath

Write-Host ''
Write-Host "Telegram subscribers: $($state.subscribers.Count)"
Write-Host ''

# ------------------------------------------------------------
# NEWS URL
# ------------------------------------------------------------

$newsUrl = [string]$config.news_url

if ([string]::IsNullOrWhiteSpace($newsUrl)) {
    $newsUrl = $DefaultNewsUrl
}

# ВАЖНО:
# Здесь должен быть обычный URL.
# Никаких [https://...](https://...) внутри строки.

Write-Host "News URL: $newsUrl"

try {

    $newsUri = [uri]$newsUrl

    if ($newsUri.Scheme -notin @('http', 'https')) {
        throw "Недопустимая схема URL: $($newsUri.Scheme)"
    }
}
catch {

    throw "Некорректный news_url: $newsUrl"
}

# ------------------------------------------------------------
# ПОЛУЧАЕМ СПИСОК СТАТЕЙ
# ------------------------------------------------------------

$articles = @()

# ============================================================
# 1. Сначала обычная HTML-страница
# ============================================================

try {

    Write-Host ''
    Write-Host '--- HTML SOURCE ---'

    $newsHtml = Get-Page $newsUrl

    Write-Host "HTML length: $($newsHtml.Length)"

    $links = Get-ArticleLinks `
        $newsHtml `
        $newsUri

    Write-Host "Article links found: $($links.Count)"

    foreach ($link in $links) {

        $articles += [pscustomobject]@{
            Url      = $link
            Title    = ''
            BodyHtml = ''
        }
    }
}
catch {

    Write-Warning `
        "HTML источник недоступен: $($_.Exception.Message)"
}

# ============================================================
# 2. RSS — только дополнительный источник
# ============================================================

$rssUrl =
    "$newsUrl?format=feed&type=rss"

try {

    Write-Host ''
    Write-Host '--- RSS SOURCE ---'
    Write-Host "RSS URL: $rssUrl"

    $rssArticles =
        Get-RssArticles $rssUrl

    Write-Host "RSS articles found: $($rssArticles.Count)"

    foreach ($rssArticle in $rssArticles) {

        $exists =
            $articles |
            Where-Object {
                $_.Url -eq $rssArticle.Url
            }

        if (-not $exists) {

            $articles +=
                $rssArticle
        }
    }
}
catch {

    Write-Warning `
        "RSS недоступен: $($_.Exception.Message)"
}

# ------------------------------------------------------------
# ЕСЛИ НИЧЕГО НЕ НАШЛИ
# ------------------------------------------------------------

if ($articles.Count -eq 0) {

    Write-Warning ''
    Write-Warning '========================================'
    Write-Warning ' НЕ УДАЛОСЬ ПОЛУЧИТЬ НОВОСТИ'
    Write-Warning '========================================'
    Write-Warning ''
    Write-Warning "URL: $newsUrl"
    Write-Warning ''
    Write-Warning 'Это уже может означать, что GitHub'
    Write-Warning 'не имеет доступа к сайту Крымэнерго.'
    Write-Warning ''

    Save-JsonFile `
        $state `
        $statePath

    exit 0
}

# ------------------------------------------------------------
# УНИКАЛЬНЫЕ ССЫЛКИ
# ------------------------------------------------------------

$articles =
    @(
        $articles |
        Group-Object Url |
        ForEach-Object {
            $_.Group[0]
        }
    )

Write-Host ''
Write-Host "Unique articles: $($articles.Count)"

# ------------------------------------------------------------
# НОВЫЕ
# ------------------------------------------------------------

$newArticles =
    @(
        $articles |
        Where-Object {
            $state.sent_urls -notcontains $_.Url
        } |
        Select-Object -First 10
    )

Write-Host "New articles: $($newArticles.Count)"

# Старые → сначала новые.
[array]::Reverse($newArticles)

# ------------------------------------------------------------
# ОБРАБОТКА
# ------------------------------------------------------------

foreach ($candidate in $newArticles) {

    Write-Host ''
    Write-Host '----------------------------------------'
    Write-Host "Processing: $($candidate.Url)"

    try {

        $article =
            Get-Article `
                $candidate.Url `
                $candidate.Title `
                $candidate.BodyHtml

        Write-Host "Title: $($article.Title)"
        Write-Host "Date: $($article.Date)"
        Write-Host "Time: $($article.Time)"
        Write-Host "Area: $($article.Area)"
        Write-Host "Text length: $($article.Text.Length)"

        # ----------------------------------------------------
        # RELEVANCE
        # ----------------------------------------------------

        $haystack =
            "$($article.Title) $($article.Text)".ToLowerInvariant()

        $isRelevant = $false

        foreach ($keyword in $RelevantKeywords) {

            if ($haystack.Contains(
                $keyword.ToLowerInvariant()
            )) {

                $isRelevant = $true

                break
            }
        }

        Write-Host "Relevant: $isRelevant"

        # ----------------------------------------------------
        # SEND
        # ----------------------------------------------------

        if (
            $isRelevant -and
            $state.subscribers.Count -gt 0 -and
            (
                $state.initialized -or
                [bool]$config.send_existing_on_first_run
            )
        ) {

            Send-TelegramArticle `
                $article `
                $config `
                ([long[]]$state.subscribers)

            Write-Host 'Telegram: SENT'
        }
        else {

            Write-Host 'Telegram: NOT SENT'
        }

        # Считаем публикацию обработанной
        # только после успешного разбора.
        $state.sent_urls += $candidate.Url
    }
    catch {

        Write-Warning ''
        Write-Warning "ОШИБКА ОБРАБОТКИ:"
        Write-Warning $candidate.Url
        Write-Warning $_.Exception.Message
    }
}

# ------------------------------------------------------------
# STATE
# ------------------------------------------------------------

$state.initialized = $true

$state.sent_urls =
    @(
        $state.sent_urls |
        Select-Object -Unique |
        Select-Object -Last 500
    )

Save-JsonFile `
    $state `
    $statePath

Write-Host ''
Write-Host '========================================'
Write-Host ' DONE'
Write-Host '========================================'
