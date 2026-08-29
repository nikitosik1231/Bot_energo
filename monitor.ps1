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

# Один запрос может ждать сайт до 90 секунд.
# Если сайт временно недоступен, статья будет запрашиваться снова.
$RequestTimeoutSeconds = 15
$RetryDelaySeconds = 5

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
    param(
        [string]$Path,
        $Default
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $Default
    }

    $Raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8

    if ([string]::IsNullOrWhiteSpace($Raw)) {
        return $Default
    }

    return ($Raw | ConvertFrom-Json)
}

function Save-JsonFile {
    param(
        $Object,
        [string]$Path
    )

    $Object | ConvertTo-Json -Depth 10 |
        Set-Content -LiteralPath $Path -Encoding UTF8
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
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    $Headers = @{
        'User-Agent' =
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36'
        'Accept' =
            'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'
        'Accept-Language' =
            'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7'
        'Cache-Control' = 'no-cache'
        'Pragma' = 'no-cache'
    }

    Write-Host ('HTTP GET: {0}' -f $Url)

    $Response = Invoke-WebRequest `
        -Uri $Url `
        -Headers $Headers `
        -MaximumRedirection 10 `
        -TimeoutSec $RequestTimeoutSeconds `
        -UseBasicParsing `
        -ErrorAction Stop

    if ([string]::IsNullOrWhiteSpace($Response.Content)) {
        throw 'Сайт вернул пустой ответ.'
    }

    Write-Host ('HTTP status: {0}' -f $Response.StatusCode)

    if ($null -ne $Response.RawContentLength) {
        Write-Host ('Downloaded bytes: {0}' -f $Response.RawContentLength)
    }

    return $Response.Content
}

function Get-PageUntilSuccess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    $Attempt = 0

    while ($true) {
        $Attempt++

        try {
            Write-Host ('Attempt #{0}: {1}' -f $Attempt, $Url)
            return Get-Page -Url $Url
        }
        catch {
            $ErrorText = $_.Exception.Message

            Write-Warning (
                'Attempt #{0} failed: {1}' -f
                $Attempt,
                $ErrorText
            )

            Write-Host (
                'Повтор через {0} сек.' -f
                $RetryDelaySeconds
            )

            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

function To-AbsoluteUrl {
    param(
        [string]$Href,
        [uri]$BaseUri
    )

    if ([string]::IsNullOrWhiteSpace($Href)) {
        return $null
    }

    $CleanHref = [System.Net.WebUtility]::HtmlDecode($Href).Trim()

    try {
        $Uri = [uri]::new($BaseUri, $CleanHref)

        if ($Uri.Scheme -notin @('http', 'https')) {
            return $null
        }

        return $Uri.AbsoluteUri
    }
    catch {
        return $null
    }
}

function Normalize-ArticleUrl {
    param(
        [string]$Url
    )

    try {
        $Uri = [uri]$Url

        if ($Uri.Scheme -notin @('http', 'https')) {
            return $null
        }

        return $Uri.GetLeftPart([System.UriPartial]::Path).TrimEnd('/')
    }
    catch {
        return $null
    }
}

function Test-ArticleUrl {
    param(
        [string]$Url
    )

    try {
        $Uri = [uri]$Url

        return (
            $Uri.AbsolutePath -match '^/about/news/[^/]+$'
        )
    }
    catch {
        return $false
    }
}

function Convert-HtmlToText {
    param(
        [string]$Html
    )

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ''
    }

    $Text = $Html

    $Text = [regex]::Replace(
        $Text,
        '<br\s*/?>',
        "`n",
        'IgnoreCase'
    )

    $Text = [regex]::Replace(
        $Text,
        '</p\s*>',
        "`n",
        'IgnoreCase'
    )

    $Text = [regex]::Replace(
        $Text,
        '</div\s*>',
        "`n",
        'IgnoreCase'
    )

    $Text = [regex]::Replace(
        $Text,
        '</li\s*>',
        "`n",
        'IgnoreCase'
    )

    $Text = [regex]::Replace(
        $Text,
        '<script\b[^>]*>.*?</script>',
        '',
        'IgnoreCase,Singleline'
    )

    $Text = [regex]::Replace(
        $Text,
        '<style\b[^>]*>.*?</style>',
        '',
        'IgnoreCase,Singleline'
    )

    $Text = $Text -replace '<[^>]+>', ' '
    $Text = [System.Net.WebUtility]::HtmlDecode($Text)

    $Text = [regex]::Replace(
        $Text,
        '[ \t]+',
        ' '
    )

    $Text = [regex]::Replace(
        $Text,
        ' *\n *',
        "`n"
    )

    return $Text.Trim()
}

function Get-ArticleLinks {
    param(
        [string]$Html,
        [uri]$BaseUri
    )

    $Result = @()

    $Seen = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    $Pattern =
        '<a\b[^>]*href\s*=\s*["''](?<href>[^"'']+)["''][^>]*>(?<text>.*?)</a>'

    $Matches = [regex]::Matches(
        $Html,
        $Pattern,
        'IgnoreCase,Singleline'
    )

    foreach ($Match in $Matches) {
        $Url = To-AbsoluteUrl `
            -Href $Match.Groups['href'].Value `
            -BaseUri $BaseUri

        if (-not $Url) {
            continue
        }

        if (-not (Test-ArticleUrl -Url $Url)) {
            continue
        }

        $Url = Normalize-ArticleUrl -Url $Url

        if (-not $Url) {
            continue
        }

        if ($Seen.Add($Url)) {
            $Title = Convert-HtmlToText `
                -Html $Match.Groups['text'].Value

            $Result += [pscustomobject]@{
                Url = $Url
                Title = $Title
            }
        }
    }

    return @($Result)
}

function Get-ArticleTitle {
    param(
        [string]$Html
    )

    $Patterns = @(
        '<h2\b[^>]*itemprop\s*=\s*["'']headline["''][^>]*>(?<value>.*?)</h2>',
        '<h1\b[^>]*>(?<value>.*?)</h1>',
        '<title\b[^>]*>(?<value>.*?)</title>'
    )

    foreach ($Pattern in $Patterns) {
        $Match = [regex]::Match(
            $Html,
            $Pattern,
            'IgnoreCase,Singleline'
        )

        if ($Match.Success) {
            $Title = Convert-HtmlToText `
                -Html $Match.Groups['value'].Value

            if ($Title) {
                return $Title
            }
        }
    }

    return 'Новости'
}

function Get-ArticleBodyHtml {
    param(
        [string]$Html
    )

    $Patterns = @(
        '<div\b[^>]*itemprop\s*=\s*["'']articleBody["''][^>]*>(?<value>.*?)</div>',
        '<article\b[^>]*>(?<value>.*?)</article>'
    )

    foreach ($Pattern in $Patterns) {
        $Match = [regex]::Match(
            $Html,
            $Pattern,
            'IgnoreCase,Singleline'
        )

        if ($Match.Success -and $Match.Groups['value'].Value) {
            return $Match.Groups['value'].Value
        }
    }

    $BodyMatch = [regex]::Match(
        $Html,
        '<body\b[^>]*>(?<value>.*?)</body>',
        'IgnoreCase,Singleline'
    )

    if ($BodyMatch.Success) {
        return $BodyMatch.Groups['value'].Value
    }

    return $Html
}

function Get-DetailLines {
    param(
        [string]$BodyHtml
    )

    $WithBreaks = $BodyHtml

    $WithBreaks = [regex]::Replace(
        $WithBreaks,
        '<br\s*/?>',
        "`n",
        'IgnoreCase'
    )

    $WithBreaks = [regex]::Replace(
        $WithBreaks,
        '</p\s*>',
        "`n",
        'IgnoreCase'
    )

    $WithBreaks = [regex]::Replace(
        $WithBreaks,
        '</div\s*>',
        "`n",
        'IgnoreCase'
    )

    $WithBreaks = [regex]::Replace(
        $WithBreaks,
        '</li\s*>',
        "`n",
        'IgnoreCase'
    )

    $Plain = Convert-HtmlToText -Html $WithBreaks
    $Result = @()

    foreach ($Raw in ($Plain -split "`r`n|`n|`r")) {
        $Line = ($Raw -replace '\s+', ' ').Trim()

        if (-not $Line) {
            continue
        }

        $IsArea = $false
        # Города и населённые пункты верхнего уровня
        if ($Line -match '^(?:г\.|город)\s*[А-ЯЁ]') {
            $IsArea = $true
        }
        if ($Line -match '^[А-ЯЁ][\p{L}\- ]+\s+район(?:а)?\s*[,;:]?$') {
            $IsArea = $true
        }
        if ($Line -match '^[А-ЯЁ][\p{L}\- ]+\s+округ(?:а)?\s*[,;:]?$') {
            $IsArea = $true
        }
            ($Line -match '^город\s+') -or
            ($Line -match '^с\.\s*') -or
            ($Line -match '^пгт\.\s*') -or
            ($Line -match 'район(?:а)?$') -or
            ($Line -match 'округ(?:а)?$')
        )

        $Result += [pscustomobject]@{
            Text = $Line.Trim(' ', ',', ':', ';')
            IsArea = $IsArea
        }
    }

    return @($Result)
}

function Get-Article {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    $Html = Get-PageUntilSuccess -Url $Url

    $Title = Get-ArticleTitle -Html $Html
    $BodyHtml = Get-ArticleBodyHtml -Html $Html
    $Text = Convert-HtmlToText -Html $BodyHtml

    $DateMatch = [regex]::Match(
        $Text,
        '\b\d{1,2}\s+(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+\d{4}\b',
        'IgnoreCase'
    )

    $Date = ''

    if ($DateMatch.Success) {
        $Date = $DateMatch.Value
    }

    $TimePeriods = @()

    $RangePattern = '(?i)(?:с\s*)?(\d{1,2}(?::\d{2})?)\s*(?:до|по|-|–|—)\s*(\d{1,2}(?::\d{2})?)'

    foreach ($Match in [regex]::Matches($Text, $RangePattern)) {
        $From = $Match.Groups[1].Value
        $To = $Match.Groups[2].Value

    if ($From -notmatch ':') {
        $From = $From + ':00'
    }

    if ($To -notmatch ':') {
        $To = $To + ':00'
    }

    $TimePeriods += "с $From до $To"
    }

    $TimePeriods = @($TimePeriods | Select-Object -Unique)

    if ($TimePeriods.Count -gt 0) {
        $Time = $TimePeriods -join ', '
    }
    else {
        $SingleTimes = @(
            [regex]::Matches($Text, '\b\d{1,2}:\d{2}\b') |
            ForEach-Object { $_.Value } |
            Select-Object -Unique
         )

        $Time = $SingleTimes -join ', '
    }

    $AreaMatch = [regex]::Match(
        $Text,
        '(?:г\.|город|с\.|пгт\.|район|округ)\s*[^,.;\n]{1,120}',
        'IgnoreCase'
    )

    $Area = ''

    if ($AreaMatch.Success) {
        $Area = $AreaMatch.Value.Trim()
    }

    $DetailLines = Get-DetailLines -BodyHtml $BodyHtml

    return [pscustomobject]@{
        Url = $Url
        Title = $Title
        Text = $Text
        Date = $Date
        Time = $Time
        Area = $Area
        DetailLines = $DetailLines
    }
}

function Test-ArticleRelevant {
    param(
        [pscustomobject]$Article
    )

    $Haystack = (
        ($Article.Title + ' ' + $Article.Text)
    ).ToLowerInvariant()

    foreach ($Keyword in $RelevantKeywords) {
        if ($Haystack.Contains(
            $Keyword.ToLowerInvariant()
        )) {
            return $true
        }
    }

    return $false
}

function Invoke-Telegram {
    param(
        [string]$Method,
        $Payload,
        $Config
    )

    $Token = [string]$Config.telegram_bot_token

    if ([string]::IsNullOrWhiteSpace($Token)) {
        throw 'telegram_bot_token не указан.'
    }

    $Json = $Payload |
        ConvertTo-Json -Compress -Depth 10

    $Uri = 'https://api.telegram.org/bot{0}/{1}' -f `
        $Token,
        $Method

    return Invoke-RestMethod `
        -Method Post `
        -Uri $Uri `
        -ContentType 'application/json; charset=utf-8' `
        -Body $Json `
        -TimeoutSec 60 `
        -ErrorAction Stop
}

function Receive-Subscribers {
    param(
        $Config,
        $State
    )

    try {
        $Response = Invoke-Telegram `
            -Method 'getUpdates' `
            -Payload @{
                offset = [long]$State.telegram_update_offset
                timeout = 0
                allowed_updates = @(
                    'message',
                    'channel_post',
                    'my_chat_member'
                )
            } `
            -Config $Config

        foreach ($Update in @($Response.result)) {
            $State.telegram_update_offset =
                [long]$Update.update_id + 1

            if ($Update.message) {
                $ChatId = [long]$Update.message.chat.id
                $Text = [string]$Update.message.text

                if ($Text -match '^/start(?:\s|$)') {
                    if ($State.subscribers -notcontains $ChatId) {
                        $State.subscribers += $ChatId
                    }

                    Invoke-Telegram `
                        -Method 'sendMessage' `
                        -Payload @{
                            chat_id = $ChatId
                            text = 'Готово. Буду присылать новые новости Крымэнерго с сообщениями об отключениях.'
                        } `
                        -Config $Config |
                        Out-Null
                }

                if ($Text -match '^/stop(?:\s|$)') {
                    $State.subscribers = @(
                        $State.subscribers |
                        Where-Object {
                            $_ -ne $ChatId
                        }
                    )

                    Invoke-Telegram `
                        -Method 'sendMessage' `
                        -Payload @{
                            chat_id = $ChatId
                            text = 'Уведомления отключены.'
                        } `
                        -Config $Config |
                        Out-Null
                }
            }

            if ($Update.channel_post) {
                $ChatId = [long]$Update.channel_post.chat.id

                if ($State.subscribers -notcontains $ChatId) {
                    $State.subscribers += $ChatId
                }
            }

            if ($Update.my_chat_member) {
                $Member = $Update.my_chat_member

                if (
                    $Member.chat.type -in @(
                        'channel',
                        'group',
                        'supergroup'
                    ) -and
                    $Member.new_chat_member.status -in @(
                        'member',
                        'administrator'
                    )
                ) {
                    $ChatId = [long]$Member.chat.id

                    if ($State.subscribers -notcontains $ChatId) {
                        $State.subscribers += $ChatId
                    }
                }
            }
        }
    }
    catch {
        Write-Warning (
            'Telegram getUpdates error: {0}' -f
            $_.Exception.Message
        )
    }
}

function Split-TelegramLine {
    param(
        [string]$Line
    )

    if ($Line.Length -le 3600) {
        return @($Line)
    }

    $Parts = @()
    $Remaining = $Line

    while ($Remaining.Length -gt 3600) {
        $Cut = $Remaining.LastIndexOf(' ', 3600)

        if ($Cut -lt 1) {
            $Cut = 3600
        }

        $Parts += $Remaining.Substring(0, $Cut)
        $Remaining =
            $Remaining.Substring($Cut).TrimStart()
    }

    if ($Remaining) {
        $Parts += $Remaining
    }

    return @($Parts)
}

function Send-TelegramLines {
    param(
        [string[]]$Lines,
        $Config,
        [long[]]$ChatIds
    )

    $Messages = @()
    $Current = ''

    foreach ($SourceLine in $Lines) {
        foreach ($Line in (
            Split-TelegramLine -Line $SourceLine
        )) {
            if ($Current) {
                $Candidate = $Current + "`n" + $Line
            }
            else {
                $Candidate = $Line
            }

            if (
                ($Candidate.Length -gt 3800) -and
                $Current
            ) {
                $Messages += $Current
                $Current =
                    "📄 <b>Продолжение</b>`n" + $Line
            }
            else {
                $Current = $Candidate
            }
        }
    }

    if ($Current) {
        $Messages += $Current
    }

    foreach ($Message in $Messages) {
        foreach ($ChatId in $ChatIds) {
            Invoke-Telegram `
                -Method 'sendMessage' `
                -Payload @{
                    chat_id = $ChatId
                    text = $Message
                    parse_mode = 'HTML'
                    disable_web_page_preview = $false
                } `
                -Config $Config |
                Out-Null
        }
    }
}

function Send-Telegram {
    param(
        [pscustomobject]$Article,
        $Config,
        [long[]]$ChatIds
    )

    # Это оформление оставлено в том же формате:
    # заголовок -> дата -> время -> территории -> ссылка.
    $Lines = @(
        '⚡ <b>Крымэнерго — отключения</b>',
        '',
        (
            '<b>📌 {0}</b>' -f
            [System.Net.WebUtility]::HtmlEncode(
                $Article.Title
            )
        )
    )

    if ($Article.Date) {
        $Lines += (
            '📅 <b>Дата:</b> {0}' -f
            [System.Net.WebUtility]::HtmlEncode(
                $Article.Date
            )
        )
    }

    if ($Article.Time) {
        $Lines += (
            '⏰ <b>Время:</b> {0}' -f
            [System.Net.WebUtility]::HtmlEncode(
                $Article.Time
            )
        )
    }

    $Lines += @(
        '',
        '📝 <b>Детали по территориям:</b>'
    )

    foreach ($Detail in $Article.DetailLines) {
        $Encoded = [System.Net.WebUtility]::HtmlEncode(
            $Detail.Text
        )

        if ($Detail.IsArea) {
            $Lines += (
                '📍 <b>' + $Encoded + '</b>'
            )
        }
        else {
            $Lines += (
                '▫️ ' + $Encoded
            )
        }
    }

    $Lines += @(
        '',
        (
            '🔗 <a href="{0}">Открыть публикацию</a>' -f
            [System.Net.WebUtility]::HtmlEncode(
                $Article.Url
            )
        )
    )

    Send-TelegramLines `
        -Lines $Lines `
        -Config $Config `
        -ChatIds $ChatIds
}

# ============================================================
# START
# ============================================================

Write-Host ''
Write-Host '========================================'
Write-Host ' Crimea Energy Telegram Monitor'
Write-Host '========================================'
Write-Host ''

$Config = Read-JsonFile `
    -Path $ConfigPath `
    -Default $null

if (-not $Config) {
    throw (
        'Не найден config.json: {0}' -f
        $ConfigPath
    )
}

if (
    [string]::IsNullOrWhiteSpace(
        [string]$Config.telegram_bot_token
    )
) {
    throw 'В config.json отсутствует telegram_bot_token.'
}

$StatePath = Join-Path `
    $ScriptDirectory `
    'state.json'

$State = Read-JsonFile `
    -Path $StatePath `
    -Default (
        [pscustomobject]@{
            known_urls = @()
            sent_urls = @()
            initialized = $false
            subscribers = @()
            telegram_update_offset = 0
        }
    )

$State = Ensure-StateProperties -State $State

Receive-Subscribers `
    -Config $Config `
    -State $State

Save-JsonFile `
    -Object $State `
    -Path $StatePath

Write-Host (
    'Telegram subscribers: {0}' -f
    @($State.subscribers).Count
)

$NewsUrl = [string]$Config.news_url

if ([string]::IsNullOrWhiteSpace($NewsUrl)) {
    $NewsUrl = $DefaultNewsUrl
}

$NewsUri = [uri]$NewsUrl

Write-Host (
    'Reading news: {0}' -f
    $NewsUrl
)

# ============================================================
# Получаем саму страницу новостей.
# Если сайт не отвечает — повторяем до успешного ответа.
# ============================================================

$NewsHtml = Get-PageUntilSuccess -Url $NewsUrl

$Links = @(Get-ArticleLinks `
    -Html $NewsHtml `
    -BaseUri $NewsUri)

Write-Host (
    'Article links found: {0}' -f
    $Links.Count
)

if ($Links.Count -eq 0) {
    Write-Warning (
        'На странице {0} не найдено ссылок на статьи.' -f
        $NewsUrl
    )

    Save-JsonFile `
        -Object $State `
        -Path $StatePath

    exit 0
}

# HashSet нужен только для быстрой проверки,
# какие URL уже были обработаны.
$KnownSet = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

foreach ($Url in @($State.known_urls)) {
    if ($Url) {
        [void]$KnownSet.Add([string]$Url)
    }
}

$SentSet = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

foreach ($Url in @($State.sent_urls)) {
    if ($Url) {
        [void]$SentSet.Add([string]$Url)
    }
}

# ============================================================
# ПЕРВЫЙ ЗАПУСК
# ============================================================

if (-not [bool]$State.initialized) {

    Write-Host ''
    Write-Host 'FIRST RUN'
    Write-Host 'Ищем 3 последние релевантные публикации.'

    $Selected = @()

    # Страница Крымэнерго отдаёт новости от новых к старым.
    # Идём сверху вниз и открываем каждую статью.
    # Если статья не открылась — Get-PageUntilSuccess
    # будет повторять запрос и НЕ даст перейти дальше.
    foreach ($Link in $Links) {

        if ($Selected.Count -ge 3) {
            break
        }

        Write-Host ''
        Write-Host (
            'Checking candidate: {0}' -f
            $Link.Url
        )

        $Article = Get-Article -Url $Link.Url

        $IsRelevant = Test-ArticleRelevant `
            -Article $Article

        Write-Host (
            'Title: {0}' -f
            $Article.Title
        )

        Write-Host (
            'Relevant: {0}' -f
            $IsRelevant
        )

        [void]$KnownSet.Add($Link.Url)

        if ($IsRelevant) {
            $Selected += $Article
        }
    }

    # Было: новая -> старая.
    # Делаем: старая -> новая.
    [array]::Reverse($Selected)

    Write-Host ''
    Write-Host (
        'FIRST RUN selected: {0}' -f
        $Selected.Count
    )

    foreach ($Article in $Selected) {

        if ($SentSet.Contains($Article.Url)) {
            continue
        }

        if (@($State.subscribers).Count -gt 0) {

            Send-Telegram `
                -Article $Article `
                -Config $Config `
                -ChatIds ([long[]]@($State.subscribers))

            Write-Host (
                'Telegram: SENT - {0}' -f
                $Article.Title
            )

            [void]$SentSet.Add($Article.Url)
        }
        else {
            Write-Warning (
                'Нет подписчиков Telegram. Не отправлена: {0}' -f
                $Article.Title
            )
        }

        [void]$KnownSet.Add($Article.Url)

        $State.known_urls = @(
            $KnownSet |
            Select-Object -Last 1000
        )

        $State.sent_urls = @(
            $SentSet |
            Select-Object -Last 500
        )

        Save-JsonFile `
            -Object $State `
            -Path $StatePath
    }

    # Всё, что было опубликовано на сайте до первого запуска,
    # считаем уже просмотренным. Поэтому на следующем запуске
    # старые статьи повторно не придут.
    foreach ($Link in $Links) {
        [void]$KnownSet.Add($Link.Url)
    }

    $State.known_urls = @(
        $KnownSet |
        Select-Object -Last 1000
    )

    $State.sent_urls = @(
        $SentSet |
        Select-Object -Last 500
    )

    $State.initialized = $true

    Save-JsonFile `
        -Object $State `
        -Path $StatePath

    Write-Host ''
    Write-Host 'FIRST RUN COMPLETE.'
    exit 0
}

# ============================================================
# ОБЫЧНЫЙ ЗАПУСК
# ============================================================

Write-Host ''
Write-Host 'NORMAL RUN'
Write-Host 'Ищем только новые публикации.'

$NewLinks = @(
    $Links |
    Where-Object {
        -not $KnownSet.Contains($_.Url)
    }
)

Write-Host (
    'New URLs: {0}' -f
    $NewLinks.Count
)

if ($NewLinks.Count -eq 0) {

    # На всякий случай запоминаем все найденные URL.
    foreach ($Link in $Links) {
        [void]$KnownSet.Add($Link.Url)
    }

    $State.known_urls = @(
        $KnownSet |
        Select-Object -Last 1000
    )

    Save-JsonFile `
        -Object $State `
        -Path $StatePath

    Write-Host 'No new publications.'
    exit 0
}

# На странице порядок: новая -> старая.
# Обрабатываем: старая новая -> новая новая.
[array]::Reverse($NewLinks)

foreach ($Link in $NewLinks) {

    Write-Host ''
    Write-Host '----------------------------------------'
    Write-Host (
        'NEW ARTICLE: {0}' -f
        $Link.Url
    )

    # К следующей статье не переходим,
    # пока текущая не будет успешно загружена.
    $Article = Get-Article -Url $Link.Url

    $IsRelevant = Test-ArticleRelevant `
        -Article $Article

    Write-Host (
        'Title: {0}' -f
        $Article.Title
    )

    Write-Host (
        'Relevant: {0}' -f
        $IsRelevant
    )

    if ($IsRelevant) {

        if (@($State.subscribers).Count -gt 0) {

            Send-Telegram `
                -Article $Article `
                -Config $Config `
                -ChatIds ([long[]]@($State.subscribers))

            Write-Host 'Telegram: SENT'

            [void]$SentSet.Add($Link.Url)
        }
        else {
            Write-Warning (
                'Нет подписчиков Telegram. Не отправлена: {0}' -f
                $Article.Title
            )
        }
    }
    else {
        Write-Host 'Telegram: NOT SENT (not relevant)'
    }

    # Независимо от релевантности статья теперь обработана.
    [void]$KnownSet.Add($Link.Url)

    $State.known_urls = @(
        $KnownSet |
        Select-Object -Last 1000
    )

    $State.sent_urls = @(
        $SentSet |
        Select-Object -Last 500
    )

    # Сохраняем состояние после каждой статьи.
    # Если GitHub оборвёт job после этого места,
    # уже обработанные статьи не будут отправлены повторно.
    Save-JsonFile `
        -Object $State `
        -Path $StatePath
}

Write-Host ''
Write-Host 'NORMAL RUN COMPLETE.'
