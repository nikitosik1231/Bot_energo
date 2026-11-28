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

# ============================================================
# НАСТРОЙКИ (можно менять)
# ============================================================

$DefaultNewsUrl = 'https://crimea-energy.ru/about/news'

# Какие типы публикаций присылать:
#   outage  - отключения
#   warning - предупреждения (погода, режим готовности и т.п.)
#   news    - обычные новости
# Чтобы не получать обычные новости, удалите слово 'news' из списка.
$SendTypes = @('outage', 'warning', 'news')

# Запросы к сайту
$RequestTimeoutSeconds = 20
$RequestAttempts = 3
$RetryDelaySeconds = 5

# Слежение за изменениями в уже отправленных публикациях
$WatchHours = 72          # сколько часов после публикации следим за правками
$FastWatchHours = 6       # первые часы проверяем при каждом запуске, дальше - примерно раз в час
$MaxWatchedPerRun = 8     # не больше стольких публикаций за один запуск
$MaxUpdatesPerArticle = 5 # не больше стольких уведомлений об изменениях по одной публикации

# Размер истории в state.json
$MaxTrackedArticles = 80
$MaxStoredLines = 120

# ============================================================
# СПРАВОЧНИКИ
# ============================================================

$MonthList = @(
    'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
    'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'
)
$MonthNames = $MonthList -join '|'

$TypeInfo = @{
    outage  = @{ Badge = '🔴'; Label = 'ОТКЛЮЧЕНИЯ' }
    warning = @{ Badge = '🟠'; Label = 'ПРЕДУПРЕЖДЕНИЕ' }
    news    = @{ Badge = '🔵'; Label = 'НОВОСТЬ' }
}

# Признаки типа по заголовку
$OutageTitleRx = 'отключ|обесточ|график\p{L}*\s+(?:отключ|ремонт|работ)|плановы\p{L}*\s+(?:работ|ремонт)|перерыв\p{L}*\s+в\s+(?:электро|энерго)снабжен|без\s+(?:света|электро\p{L}+)|ограничени\p{L}*\s+(?:электро|энерго)'
$WarningTitleRx = 'опасн|предупрежд|гидрометеор|метеор|штормов|непогод|ураган|гололед|гололёд|снегопад|ливн|мороз|жара|пожароопасн|режим\p{L}*\s+повышенной|чрезвычайн|внимани|напомина|рекомендац|как\s+действовать|куда\s+сообщ'

# Признаки типа по тексту
$OutageBodyPatterns = @(
    'отключен', 'обесточ', 'плановы\p{L}*\s+(?:работ|ремонт)',
    'график', 'ремонтн\p{L}*\s+работ', 'перерыв\p{L}*\s+в\s+(?:электро|энерго)снабжен'
)
$WarningBodyPatterns = @(
    'порыв', 'ветр', 'шторм', 'ураган', 'гололед', 'гололёд', 'снегопад', 'ливн',
    'режим\p{L}*\s+повышенной\s+готовности', 'напоминаем', 'горяч\p{L}+\s+лини',
    'чрезвычайн', 'неблагоприятн', 'мчс', 'опасн\p{L}+\s+(?:явлен|погод)'
)

# Строки, которые не нужны в тексте публикации
$NoiseLineRx = '^(?i)(?:поделиться|просмотр\p{L}*:?\s*\d*|читайте также.*|назад|печать|версия для печати)$'

# После чисел это НЕ время (ветер, проценты, даты и т.п.)
$NotTimeTailRx = '^\s*(?:м/с|км|%|°|мм|см|м\b|градус|метр|процент|тыс|млн|руб|₽|раз|дн\p{L}*|сут\p{L}*|мин\p{L}*|недел\p{L}*|человек|лет\b|год|числа|число|(?:' + $MonthNames + '))'

# ============================================================
# JSON / СОСТОЯНИЕ
# ============================================================

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

    # Новое: журнал отправленных публикаций (для отслеживания изменений)
    if ($null -eq $State.articles) {
        $State | Add-Member -NotePropertyName articles -NotePropertyValue @()
    }

    return $State
}

# ============================================================
# ЗАГРУЗКА СТРАНИЦ
# ============================================================

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

    return [string]$Response.Content
}

# Несколько попыток. Если сайт не отвечает - возвращает $null
# (бот просто попробует в следующий запуск, а не будет висеть).
function Get-PageWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    $script:LastHttpStatus = 0

    for ($Attempt = 1; $Attempt -le $RequestAttempts; $Attempt++) {
        try {
            Write-Host ('HTTP GET ({0}/{1}): {2}' -f $Attempt, $RequestAttempts, $Url)
            return (Get-Page -Url $Url)
        }
        catch {
            $Code = 0

            if ($_.Exception.Response) {
                try { $Code = [int]$_.Exception.Response.StatusCode } catch { $Code = 0 }
            }

            $script:LastHttpStatus = $Code

            Write-Warning ('Попытка {0} не удалась: {1}' -f $Attempt, $_.Exception.Message)

            # Страницы нет совсем - повторять бессмысленно
            if ($Code -in @(404, 410)) {
                return $null
            }

            if ($Attempt -lt $RequestAttempts) {
                Start-Sleep -Seconds $RetryDelaySeconds
            }
        }
    }

    return $null
}

# ============================================================
# ССЫЛКИ
# ============================================================

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
        return ($Uri.AbsolutePath -match '^/about/news/[^/]+$')
    }
    catch {
        return $false
    }
}

# ============================================================
# РАЗБОР HTML
# ============================================================

function Convert-HtmlToText {
    param(
        [string]$Html
    )

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ''
    }

    $Text = $Html

    $Text = [regex]::Replace($Text, '<script\b[^>]*>.*?</script>', '', 'IgnoreCase,Singleline')
    $Text = [regex]::Replace($Text, '<style\b[^>]*>.*?</style>', '', 'IgnoreCase,Singleline')
    $Text = [regex]::Replace($Text, '<noscript\b[^>]*>.*?</noscript>', '', 'IgnoreCase,Singleline')

    $Text = [regex]::Replace($Text, '<br\s*/?>', "`n", 'IgnoreCase')
    $Text = [regex]::Replace($Text, '</(?:p|div|li|tr|h[1-6]|ul|ol|table|section)\s*>', "`n", 'IgnoreCase')
    $Text = [regex]::Replace($Text, '<li\b[^>]*>', "`n", 'IgnoreCase')

    $Text = $Text -replace '<[^>]+>', ' '
    $Text = [System.Net.WebUtility]::HtmlDecode($Text)
    $Text = $Text.Replace([string][char]0x00A0, ' ')

    $Text = [regex]::Replace($Text, '[ \t]+', ' ')
    $Text = [regex]::Replace($Text, ' *\n *', "`n")
    $Text = [regex]::Replace($Text, '\n{3,}', "`n`n")

    return $Text.Trim()
}

# <div ...> ... </div> с учётом вложенных div (раньше текст обрезался на первом </div>)
function Get-BalancedDivContent {
    param(
        [string]$Html,
        [string]$OpenPattern
    )

    $Open = [regex]::Match($Html, $OpenPattern, 'IgnoreCase,Singleline')

    if (-not $Open.Success) {
        return $null
    }

    $Start = $Open.Index + $Open.Length
    $Rest = $Html.Substring($Start)
    $Depth = 1

    foreach ($Tag in [regex]::Matches($Rest, '<(/?)div\b[^>]*>', 'IgnoreCase')) {
        if ($Tag.Groups[1].Value -eq '/') {
            $Depth--
        }
        else {
            $Depth++
        }

        if ($Depth -eq 0) {
            return $Rest.Substring(0, $Tag.Index)
        }
    }

    return $Rest
}

function Get-ArticleBodyHtml {
    param(
        [string]$Html
    )

    $Div = Get-BalancedDivContent `
        -Html $Html `
        -OpenPattern '<div\b[^>]*itemprop\s*=\s*["'']articleBody["''][^>]*>'

    if ($Div) {
        return $Div
    }

    $Article = [regex]::Match($Html, '<article\b[^>]*>(?<value>.*?)</article>', 'IgnoreCase,Singleline')

    if ($Article.Success -and $Article.Groups['value'].Value) {
        return $Article.Groups['value'].Value
    }

    $Body = [regex]::Match($Html, '<body\b[^>]*>(?<value>.*?)</body>', 'IgnoreCase,Singleline')

    if ($Body.Success) {
        return $Body.Groups['value'].Value
    }

    return $Html
}

function Get-ArticleLinks {
    param(
        [string]$Html,
        [uri]$BaseUri
    )

    # url -> лучший (самый длинный) текст ссылки; порядок как на странице
    $Found = [ordered]@{}

    $Pattern = '<a\b[^>]*href\s*=\s*["''](?<href>[^"'']+)["''][^>]*>(?<text>.*?)</a>'

    foreach ($Match in [regex]::Matches($Html, $Pattern, 'IgnoreCase,Singleline')) {
        $Url = To-AbsoluteUrl -Href $Match.Groups['href'].Value -BaseUri $BaseUri

        if (-not $Url) { continue }
        if (-not (Test-ArticleUrl -Url $Url)) { continue }

        $Url = Normalize-ArticleUrl -Url $Url

        if (-not $Url) { continue }

        $Title = Convert-HtmlToText -Html $Match.Groups['text'].Value

        if (-not $Found.Contains($Url)) {
            $Found[$Url] = $Title
        }
        elseif ($Title.Length -gt ([string]$Found[$Url]).Length) {
            $Found[$Url] = $Title
        }
    }

    $Result = @()

    foreach ($Key in $Found.Keys) {
        $Result += [pscustomobject]@{
            Url = [string]$Key
            Title = [string]$Found[$Key]
        }
    }

    return @($Result)
}

# ============================================================
# ЗАГОЛОВОК
# ============================================================

function Clean-Title {
    param(
        [string]$Title
    )

    if ([string]::IsNullOrWhiteSpace($Title)) {
        return ''
    }

    $T = [System.Net.WebUtility]::HtmlDecode($Title)
    $T = ($T -replace '\s+', ' ').Trim()

    # Убираем хвост с названием сайта: "... | Крымэнерго", "... - Новости"
    $T = $T -replace '\s*[|\-–—]\s*(?:ГУП\s+РК\s*)?[«"]?Крымэнерго[»"]?\s*$', ''
    $T = $T -replace '\s*[|\-–—]\s*Новости\s*$', ''
    $T = $T.Trim(' ', '|', '-', '–', '—')

    # Слишком общие заголовки не годятся
    if ($T -match '^(?i)(?:новости|крымэнерго|главная|news|без названия)$') {
        return ''
    }

    if ($T.Length -lt 4) {
        return ''
    }

    if ($T.Length -gt 220) {
        $T = $T.Substring(0, 217).TrimEnd() + '…'
    }

    return $T
}

function Get-ArticleTitle {
    param(
        [string]$Html,
        [string]$FallbackTitle,
        [string[]]$BodyLines
    )

    $Candidates = @()

    $Patterns = @(
        '<h[12]\b[^>]*itemprop\s*=\s*["'']headline["''][^>]*>(?<value>.*?)</h[12]>',
        '<meta\b[^>]*property\s*=\s*["'']og:title["''][^>]*content\s*=\s*["''](?<value>[^"'']*)["'']',
        '<meta\b[^>]*content\s*=\s*["''](?<value>[^"'']*)["''][^>]*property\s*=\s*["'']og:title["'']',
        '<h1\b[^>]*>(?<value>.*?)</h1>',
        '<title\b[^>]*>(?<value>.*?)</title>'
    )

    foreach ($Pattern in $Patterns) {
        $Match = [regex]::Match($Html, $Pattern, 'IgnoreCase,Singleline')

        if ($Match.Success) {
            $Candidates += (Convert-HtmlToText -Html $Match.Groups['value'].Value)
        }
    }

    # Название со страницы со списком новостей
    $Candidates += $FallbackTitle

    foreach ($Candidate in $Candidates) {
        $Clean = Clean-Title -Title $Candidate

        if ($Clean) {
            return $Clean
        }
    }

    # Крайний случай: первая строка текста
    if (@($BodyLines).Count -gt 0) {
        $First = [string]@($BodyLines)[0]

        if ($First.Length -gt 120) {
            $First = $First.Substring(0, 117).TrimEnd() + '…'
        }

        return $First
    }

    return 'Публикация Крымэнерго'
}

# ============================================================
# ТЕКСТ ПУБЛИКАЦИИ
# ============================================================

function Get-NormalizedKey {
    param(
        [string]$Text
    )

    if ($null -eq $Text) {
        return ''
    }

    return ($Text.ToLowerInvariant() -replace '[^\p{L}\p{Nd}]+', '')
}

function Get-Sha256 {
    param(
        [string]$Text
    )

    $Sha = [System.Security.Cryptography.SHA256]::Create()

    try {
        $Bytes = $Sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Text))
        return [System.BitConverter]::ToString($Bytes).Replace('-', '')
    }
    finally {
        $Sha.Dispose()
    }
}

function Get-BodyLines {
    param(
        [string]$BodyHtml,
        [string]$Title
    )

    $Plain = Convert-HtmlToText -Html $BodyHtml
    $TitleKey = Get-NormalizedKey -Text $Title
    $Result = New-Object System.Collections.Generic.List[string]
    $PrevKey = ''

    foreach ($Raw in ($Plain -split "`r`n|`n|`r")) {
        $Line = ($Raw -replace '\s+', ' ').Trim()

        if (-not $Line) { continue }
        if ($Line -match $NoiseLineRx) { continue }

        $Key = Get-NormalizedKey -Text $Line

        if (-not $Key) { continue }

        # Заголовок, повторённый в начале текста
        if (($Key -eq $TitleKey) -and ($Result.Count -eq 0)) { continue }

        # Подряд идущие дубли
        if ($Key -eq $PrevKey) { continue }

        $PrevKey = $Key
        $Result.Add($Line)
    }

    return @($Result)
}

# Ссылки из текста и из <a href> - в кнопки под сообщением
function Get-ExtraLinks {
    param(
        [string]$BodyHtml,
        [string]$PlainText,
        [string]$ArticleUrl
    )

    $OwnUrl = Normalize-ArticleUrl -Url $ArticleUrl
    $Seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $Result = @()
    $All = @()

    foreach ($Match in [regex]::Matches($BodyHtml, '<a\b[^>]*href\s*=\s*["''](?<href>[^"'']+)["'']', 'IgnoreCase,Singleline')) {
        $All += [System.Net.WebUtility]::HtmlDecode($Match.Groups['href'].Value).Trim()
    }

    foreach ($Match in [regex]::Matches($PlainText, 'https?://[^\s<>"'']+')) {
        $All += $Match.Value.TrimEnd('.', ',', ';', ')', ':')
    }

    foreach ($Candidate in $All) {
        if ($Candidate -notmatch '^https?://') { continue }

        try { $Uri = [uri]$Candidate } catch { continue }

        $Normalized = $Candidate.TrimEnd('/')

        if ($OwnUrl -and ($Normalized -eq $OwnUrl)) { continue }
        if (-not $Seen.Add($Normalized)) { continue }

        $Result += [pscustomobject]@{
            Url = $Candidate
            Label = (Get-LinkLabel -Uri $Uri)
        }
    }

    return @($Result | Select-Object -First 6)
}

function Get-LinkLabel {
    param(
        [uri]$Uri
    )

    $HostName = ($Uri.Host.ToLowerInvariant() -replace '^www\.', '')

    if ($HostName -match '(^|\.)max\.ru$') { return 'MAX' }
    if ($HostName -match '(^|\.)vk\.(com|ru)$') { return 'ВКонтакте' }
    if ($HostName -match '(^|\.)ok\.ru$') { return 'Одноклассники' }
    if ($HostName -match '(^|\.)t\.me$') { return 'Telegram' }
    if ($HostName -match '(^|\.)crimea-energy\.ru$') { return 'Ещё на сайте' }

    return $HostName
}

# Убираем ссылки из строк (они уходят в кнопки)
function Remove-UrlsFromLines {
    param(
        [string[]]$Lines
    )

    $Result = @()

    foreach ($Line in $Lines) {
        if ($Line -notmatch 'https?://') {
            $Result += $Line
            continue
        }

        $Clean = ($Line -replace 'https?://\S+', '').Trim()
        $Clean = ($Clean -replace '\s+', ' ').Trim(' ', '-', '–', '—')

        if ($Clean.Length -lt 3) { continue }
        if ($Clean -match '^[\p{P}\s]*$') { continue }
        if ($Clean.EndsWith(':')) { continue }

        $Result += $Clean
    }

    return @($Result)
}

# ============================================================
# ДАТЫ И ВРЕМЯ
# ============================================================

function Get-DateMentions {
    param(
        [string]$Text
    )

    $Items = @()

    $P1 = '(?<![\d.])(?<d>\d{1,2})\s+(?<m>' + $MonthNames + ')(?:\s+(?<y>\d{4}))?(?!\d)'

    foreach ($Match in [regex]::Matches($Text, $P1, 'IgnoreCase')) {
        $Day = [int]$Match.Groups['d'].Value

        if ($Day -lt 1 -or $Day -gt 31) { continue }

        $Items += [pscustomobject]@{
            D = $Day
            M = $Match.Groups['m'].Value.ToLowerInvariant()
            Y = $Match.Groups['y'].Value
        }
    }

    foreach ($Match in [regex]::Matches($Text, '(?<![\d.])(?<d>\d{1,2})\.(?<mo>\d{2})\.(?<y>\d{4})(?!\d)')) {
        $Day = [int]$Match.Groups['d'].Value
        $Mon = [int]$Match.Groups['mo'].Value

        if ($Day -lt 1 -or $Day -gt 31) { continue }
        if ($Mon -lt 1 -or $Mon -gt 12) { continue }

        $Items += [pscustomobject]@{
            D = $Day
            M = $MonthList[$Mon - 1]
            Y = $Match.Groups['y'].Value
        }
    }

    # Если год указан у всех дат один и тот же - дописываем его датам без года
    $Years = @($Items | Where-Object { $_.Y } | ForEach-Object { $_.Y } | Select-Object -Unique)

    $Out = @()

    foreach ($Item in $Items) {
        $Y = $Item.Y

        if ((-not $Y) -and ($Years.Count -eq 1)) {
            $Y = $Years[0]
        }

        $Text2 = '{0} {1}' -f $Item.D, $Item.M

        if ($Y) {
            $Text2 += ' ' + $Y
        }

        if ($Out -notcontains $Text2) {
            $Out += $Text2
        }
    }

    return @($Out | Select-Object -First 3)
}

# Время отключения: только настоящие периоды "с 09:00 до 17:00".
# Скорость ветра "15-20 м/с", даты, проценты и несуществующее время вроде 25:00 отбрасываются.
function Get-TimePeriods {
    param(
        [string]$Text
    )

    $Result = @()

    $HourMark = '(?:\s*ч(?:ас\p{L}*)?\.?)'

    $PatternA = '(?<![\p{L}\d:.,])[сc]\s*(?<h1>\d{1,2})(?:[:.](?<m1>\d{2}))?(?<hr1>' + $HourMark + ')?\s*(?:до|по|-|–|—)\s*(?<h2>\d{1,2})(?:[:.](?<m2>\d{2}))?(?<hr2>' + $HourMark + ')?(?!\d)(?![.:]\d)'
    $PatternB = '(?<![\d:.,])(?<h1>\d{1,2}):(?<m1>\d{2})\s*(?:до|по|-|–|—)\s*(?<h2>\d{1,2}):(?<m2>\d{2})(?![\d:])'

    foreach ($Pattern in @($PatternA, $PatternB)) {
        foreach ($Match in [regex]::Matches($Text, $Pattern, 'IgnoreCase')) {
            $H1 = [int]$Match.Groups['h1'].Value
            $H2 = [int]$Match.Groups['h2'].Value
            $M1 = 0
            $M2 = 0
            $HasMinutes = $false

            if ($Match.Groups['m1'].Success -and $Match.Groups['m2'].Success) {
                $HasMinutes = $true
            }

            if ($Match.Groups['m1'].Success) { $M1 = [int]$Match.Groups['m1'].Value }
            if ($Match.Groups['m2'].Success) { $M2 = [int]$Match.Groups['m2'].Value }

            # Корректное время
            if ($H1 -gt 23) { continue }
            if ($H2 -gt 24) { continue }
            if ($M1 -gt 59 -or $M2 -gt 59) { continue }
            if (($H2 -eq 24) -and ($M2 -ne 0)) { continue }
            if (($H1 -eq $H2) -and ($M1 -eq $M2)) { continue }

            # Что идёт сразу после (м/с, %, месяц и т.п.) - значит это не время
            $TailStart = $Match.Index + $Match.Length
            $Tail = $Text.Substring($TailStart, [Math]::Min(20, $Text.Length - $TailStart))

            if ($Tail -match $NotTimeTailRx) { continue }

            # "с 9 до 17" без минут принимаем только если рядом "час" или конец фразы
            if ((-not $HasMinutes) -and ($Pattern -eq $PatternA)) {
                $HasHourWord = $Match.Groups['hr1'].Success -or $Match.Groups['hr2'].Success
                $EndsPhrase = ($Tail -match '^\s*(?:[,.;:)\n]|$)')

                if ((-not $HasHourWord) -and (-not $EndsPhrase)) { continue }
            }

            $Period = 'с {0:00}:{1:00} до {2:00}:{3:00}' -f $H1, $M1, $H2, $M2

            if ($Result -notcontains $Period) {
                $Result += $Period
            }
        }
    }

    return @($Result)
}

# ============================================================
# ТИП ПУБЛИКАЦИИ
# ============================================================

function Get-ArticleType {
    param(
        [string]$Title,
        [string]$Text,
        [string[]]$TimePeriods,
        [string[]]$Lines
    )

    $TitleLower = $Title.ToLowerInvariant()
    $BodyLower = $Text.ToLowerInvariant()

    # 1. По заголовку: он самый надёжный
    if ($TitleLower -match $OutageTitleRx) { return 'outage' }
    if ($TitleLower -match $WarningTitleRx) { return 'warning' }

    # 2. По тексту: считаем баллы
    $OutageScore = 0

    foreach ($Pattern in $OutageBodyPatterns) {
        if ($BodyLower -match $Pattern) { $OutageScore++ }
    }

    if (@($TimePeriods).Count -gt 0) { $OutageScore += 3 }

    $AreaCount = @($Lines | Where-Object { Test-AreaLine -Line $_ }).Count

    if ($AreaCount -ge 2) { $OutageScore++ }

    $WarningScore = 0

    foreach ($Pattern in $WarningBodyPatterns) {
        if ($BodyLower -match $Pattern) { $WarningScore++ }
    }

    if (($OutageScore -ge 3) -and ($OutageScore -ge $WarningScore)) { return 'outage' }
    if ($WarningScore -ge 2) { return 'warning' }

    return 'news'
}

function Test-AreaLine {
    param(
        [string]$Line
    )

    if ($Line.Length -gt 90) {
        return $false
    }

    # Строка со временем - это уже описание, а не название территории
    if ($Line -match '\d{1,2}[:.]\d{2}') {
        return $false
    }

    if ($Line -match '^(?:г\.|г\s|город\s|с\.|пгт\.?|пос\.|посёлок|поселок)\s*\p{Lu}') {
        return $true
    }

    if ($Line -match '(?:район\p{L}*|округ\p{L}*)\s*[,;:.]?$') {
        return $true
    }

    return $false
}

# ============================================================
# СТАТЬЯ
# ============================================================

function Get-Article {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url,
        [string]$FallbackTitle
    )

    $Html = Get-PageWithRetry -Url $Url

    if (-not $Html) {
        return $null
    }

    $BodyHtml = Get-ArticleBodyHtml -Html $Html
    $RawLines = @(Get-BodyLines -BodyHtml $BodyHtml -Title '')
    $Title = Get-ArticleTitle -Html $Html -FallbackTitle $FallbackTitle -BodyLines $RawLines

    # Повторно собираем строки уже зная заголовок (чтобы убрать его дубль в начале текста)
    $RawLines = @(Get-BodyLines -BodyHtml $BodyHtml -Title $Title)

    $PlainText = ($RawLines -join "`n")
    $Links = @(Get-ExtraLinks -BodyHtml $BodyHtml -PlainText $PlainText -ArticleUrl $Url)

    $Lines = @(Remove-UrlsFromLines -Lines $RawLines | Select-Object -First $MaxStoredLines)
    $Text = ($Lines -join "`n")

    $TimePeriods = @(Get-TimePeriods -Text $Text)
    $Dates = @(Get-DateMentions -Text ($Title + "`n" + $Text))
    $Type = Get-ArticleType -Title $Title -Text ($Title + "`n" + $Text) -TimePeriods $TimePeriods -Lines $Lines

    $HashSource = (Get-NormalizedKey -Text $Title) + '|' +
        (($Lines | ForEach-Object { Get-NormalizedKey -Text $_ }) -join '|')

    return [pscustomobject]@{
        Url = $Url
        Title = $Title
        Type = $Type
        Date = ($Dates -join ', ')
        Time = ($TimePeriods -join ', ')
        Lines = $Lines
        Links = $Links
        Hash = (Get-Sha256 -Text $HashSource)
    }
}

# ============================================================
# TELEGRAM: ОТПРАВКА
# ============================================================

function ConvertTo-TgHtml {
    param(
        [string]$Text
    )

    # Telegram требует экранировать только эти три символа
    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
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

    $Json = $Payload | ConvertTo-Json -Compress -Depth 10

    $Uri = 'https://api.telegram.org/bot{0}/{1}' -f $Token, $Method

    for ($Try = 1; $Try -le 3; $Try++) {
        try {
            return Invoke-RestMethod `
                -Method Post `
                -Uri $Uri `
                -ContentType 'application/json; charset=utf-8' `
                -Body $Json `
                -TimeoutSec 60 `
                -ErrorAction Stop
        }
        catch {
            $Details = [string]$_.ErrorDetails.Message

            # Telegram просит притормозить
            if (($Details -match '"error_code":429') -and ($Try -lt 3)) {
                $Wait = 5

                if ($Details -match '"retry_after":(\d+)') {
                    $Wait = [int]$Matches[1] + 1
                }

                Write-Warning ('Telegram: лимит, ждём {0} сек.' -f $Wait)
                Start-Sleep -Seconds $Wait
                continue
            }

            throw
        }
    }
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
            $State.telegram_update_offset = [long]$Update.update_id + 1

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
                            text = 'Готово. Буду присылать новости Крымэнерго: отключения, предупреждения и другие публикации. Если новость изменится, отвечу на неё с описанием правок. Отписаться: /stop'
                        } `
                        -Config $Config |
                        Out-Null
                }

                if ($Text -match '^/stop(?:\s|$)') {
                    $State.subscribers = @(
                        $State.subscribers |
                        Where-Object { $_ -ne $ChatId }
                    )

                    Invoke-Telegram `
                        -Method 'sendMessage' `
                        -Payload @{
                            chat_id = $ChatId
                            text = 'Уведомления отключены. Вернуть: /start'
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
                    $Member.chat.type -in @('channel', 'group', 'supergroup') -and
                    $Member.new_chat_member.status -in @('member', 'administrator')
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
        Write-Warning ('Telegram getUpdates error: {0}' -f $_.Exception.Message)
    }
}

# Разбивка слишком длинной строки
function Split-LongLine {
    param(
        [string]$Line,
        [int]$Max = 2500
    )

    if ($Line.Length -le $Max) {
        return @($Line)
    }

    $Parts = @()
    $Remaining = $Line

    while ($Remaining.Length -gt $Max) {
        $Cut = $Remaining.LastIndexOf(' ', $Max)

        if ($Cut -lt 1) {
            $Cut = $Max
        }

        $Parts += $Remaining.Substring(0, $Cut)
        $Remaining = $Remaining.Substring($Cut).TrimStart()
    }

    if ($Remaining) {
        $Parts += $Remaining
    }

    return @($Parts)
}

function Get-Buttons {
    param(
        $Article,
        [switch]$SiteOnly
    )

    $Rows = @()

    $Rows += , @(
        @{ text = '🌐 Открыть на сайте'; url = $Article.Url }
    )

    if (-not $SiteOnly) {
        $Row = @()

        foreach ($Link in @($Article.Links)) {
            $Row += @{ text = ('🔗 {0}' -f $Link.Label); url = $Link.Url }

            if ($Row.Count -eq 2) {
                $Rows += , @($Row)
                $Row = @()
            }
        }

        if ($Row.Count -gt 0) {
            $Rows += , @($Row)
        }
    }

    return , $Rows
}

# Отправка набора сообщений во все чаты.
# ReplyTo: @{ "<chat_id>" = <message_id> } - на какое сообщение отвечать.
# Возвращает @{ "<chat_id>" = <id первого сообщения> } для чатов, куда дошло.
function Send-Parts {
    param(
        [string[]]$Parts,
        $Buttons,
        $Config,
        $State,
        [long[]]$ChatIds,
        $ReplyTo
    )

    $Sent = @{}

    foreach ($ChatId in $ChatIds) {
        $FirstId = $null

        try {
            for ($i = 0; $i -lt $Parts.Count; $i++) {
                $Payload = @{
                    chat_id = $ChatId
                    text = $Parts[$i]
                    parse_mode = 'HTML'
                    link_preview_options = @{ is_disabled = $true }
                }

                if (($i -eq 0) -and $ReplyTo -and $ReplyTo.ContainsKey([string]$ChatId)) {
                    $Payload.reply_parameters = @{
                        message_id = [long]$ReplyTo[[string]$ChatId]
                        allow_sending_without_reply = $true
                    }
                }

                if (($i -eq ($Parts.Count - 1)) -and $Buttons) {
                    $Payload.reply_markup = @{ inline_keyboard = $Buttons }
                }

                try {
                    $Result = Invoke-Telegram -Method 'sendMessage' -Payload $Payload -Config $Config
                }
                catch {
                    $Details = [string]$_.ErrorDetails.Message

                    # Если Telegram не принял разметку - отправляем простым текстом
                    if ($Details -match "can't parse entities") {
                        Write-Warning 'Telegram не принял HTML-разметку, отправляю простым текстом.'

                        $Plain = [System.Net.WebUtility]::HtmlDecode(($Parts[$i] -replace '<[^>]+>', ''))
                        $Payload.text = $Plain
                        $Payload.Remove('parse_mode')

                        $Result = Invoke-Telegram -Method 'sendMessage' -Payload $Payload -Config $Config
                    }
                    else {
                        throw
                    }
                }

                if (($i -eq 0) -and $Result.result.message_id) {
                    $FirstId = [long]$Result.result.message_id
                }

                Start-Sleep -Milliseconds 150
            }

            if ($null -ne $FirstId) {
                $Sent[[string]$ChatId] = $FirstId
            }
        }
        catch {
            $Details = [string]$_.ErrorDetails.Message

            Write-Warning ('Не удалось отправить в чат {0}: {1} {2}' -f $ChatId, $_.Exception.Message, $Details)

            # Бота заблокировали или выгнали из чата - больше туда не пишем
            if ($Details -match '"error_code":403') {
                $State.subscribers = @(
                    $State.subscribers |
                    Where-Object { $_ -ne $ChatId }
                )

                Write-Warning ('Чат {0} удалён из подписчиков.' -f $ChatId)
            }
        }
    }

    return $Sent
}

# ============================================================
# TELEGRAM: ОФОРМЛЕНИЕ
# ============================================================

function Format-BodyLine {
    param(
        [string]$Line,
        [string]$Type
    )

    $Encoded = ConvertTo-TgHtml -Text $Line

    if (($Type -eq 'outage') -and (Test-AreaLine -Line $Line)) {
        return ('📍 <b>' + $Encoded + '</b>')
    }

    # Короткие строки ЗАГЛАВНЫМИ (например "НАПОМИНАЕМ") выделяем жирным
    if (($Line.Length -le 40) -and ($Line -cmatch '^[^\p{Ll}]*\p{Lu}[^\p{Ll}]*$')) {
        return ('<b>' + $Encoded + '</b>')
    }

    return $Encoded
}

function Join-Quote {
    param(
        [string]$Prefix,
        [string[]]$Lines,
        [bool]$Expandable
    )

    if (@($Lines).Count -eq 0) {
        return $Prefix
    }

    $Tag = '<blockquote>'

    if ($Expandable) {
        $Tag = '<blockquote expandable>'
    }

    return ($Prefix + "`n`n" + $Tag + ($Lines -join "`n") + '</blockquote>')
}

function Format-PostParts {
    param(
        $Article
    )

    $Info = $TypeInfo[$Article.Type]
    $Limit = 3600

    $Head = @()
    $Head += ('{0} <b>{1}</b> · Крымэнерго' -f $Info.Badge, $Info.Label)
    $Head += ''
    $Head += ('<b>{0}</b>' -f (ConvertTo-TgHtml -Text $Article.Title))

    $Meta = @()

    if (($Article.Type -in @('outage', 'warning')) -and $Article.Date) {
        $Meta += ('📅 {0}' -f (ConvertTo-TgHtml -Text $Article.Date))
    }

    if (($Article.Type -eq 'outage') -and $Article.Time) {
        $Meta += ('🕒 {0}' -f (ConvertTo-TgHtml -Text $Article.Time))
    }

    if ($Meta.Count -gt 0) {
        $Head += ''
        $Head += $Meta
    }

    $HeadText = ($Head -join "`n")

    $BodyLines = @()

    foreach ($Line in @($Article.Lines)) {
        foreach ($Piece in (Split-LongLine -Line $Line)) {
            $BodyLines += (Format-BodyLine -Line $Piece -Type $Article.Type)
        }
    }

    $Total = 0
    foreach ($B in $BodyLines) { $Total += $B.Length + 1 }

    $Expandable = ($Total -gt 600)

    $Parts = @()
    $Prefix = $HeadText
    $Current = @()
    $CurrentLen = 0

    foreach ($B in $BodyLines) {
        if (
            ($Current.Count -gt 0) -and
            (($Prefix.Length + $CurrentLen + $B.Length + 80) -gt $Limit)
        ) {
            $Parts += (Join-Quote -Prefix $Prefix -Lines $Current -Expandable $Expandable)
            $Prefix = '📄 <i>Продолжение</i>'
            $Current = @()
            $CurrentLen = 0
        }

        $Current += $B
        $CurrentLen += $B.Length + 1
    }

    $Parts += (Join-Quote -Prefix $Prefix -Lines $Current -Expandable $Expandable)

    return @($Parts)
}

# ============================================================
# ИЗМЕНЕНИЯ В ПУБЛИКАЦИИ
# ============================================================

function Compare-Article {
    param(
        $Record,
        $Article
    )

    $OldKeys = [System.Collections.Generic.HashSet[string]]::new()
    $NewKeys = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($Line in @($Record.lines)) {
        [void]$OldKeys.Add((Get-NormalizedKey -Text ([string]$Line)))
    }

    foreach ($Line in @($Article.Lines)) {
        [void]$NewKeys.Add((Get-NormalizedKey -Text ([string]$Line)))
    }

    $Added = @($Article.Lines | Where-Object { -not $OldKeys.Contains((Get-NormalizedKey -Text $_)) })
    $Removed = @($Record.lines | Where-Object { -not $NewKeys.Contains((Get-NormalizedKey -Text ([string]$_))) })

    $TitleChanged = (
        (Get-NormalizedKey -Text ([string]$Record.title)) -ne
        (Get-NormalizedKey -Text $Article.Title)
    )

    $OldTime = [string]$Record.time
    $OldDate = [string]$Record.date

    $TimeChanged = ($Article.Type -eq 'outage') -and ($OldTime -ne $Article.Time) -and ($OldTime -or $Article.Time)
    $DateChanged = ($Article.Type -in @('outage', 'warning')) -and ($OldDate -ne $Article.Date) -and ($OldDate -or $Article.Date)

    return [pscustomobject]@{
        HasChanges = ($TitleChanged -or ($Added.Count -gt 0) -or ($Removed.Count -gt 0))
        TitleChanged = $TitleChanged
        OldTitle = [string]$Record.title
        Added = $Added
        Removed = $Removed
        TimeChanged = $TimeChanged
        OldTime = $OldTime
        DateChanged = $DateChanged
        OldDate = $OldDate
    }
}

function Format-ShortLines {
    param(
        [string[]]$Lines,
        [int]$Max = 5
    )

    $Out = @()
    $Count = 0

    foreach ($Line in $Lines) {
        if ($Count -ge $Max) {
            $Out += ('… и ещё {0}' -f (@($Lines).Count - $Max))
            break
        }

        $Text = $Line

        if ($Text.Length -gt 250) {
            $Text = $Text.Substring(0, 247).TrimEnd() + '…'
        }

        $Out += (ConvertTo-TgHtml -Text $Text)
        $Count++
    }

    return $Out
}

function Format-ChangeMessage {
    param(
        $Article,
        $Diff
    )

    $Info = $TypeInfo[$Article.Type]

    $Text = @()
    $Text += ('🔄 <b>Публикация обновлена</b> · {0}' -f $Info.Label)
    $Text += ('<b>{0}</b>' -f (ConvertTo-TgHtml -Text $Article.Title))
    $Text += ''
    $Text += '✏️ <b>Что изменилось</b>'

    if ($Diff.TitleChanged) {
        $Text += ''
        $Text += ('Заголовок: <s>{0}</s> → {1}' -f
            (ConvertTo-TgHtml -Text $Diff.OldTitle),
            (ConvertTo-TgHtml -Text $Article.Title))
    }

    if ($Diff.DateChanged) {
        $Text += ''
        $Text += ('📅 Дата: <s>{0}</s> → {1}' -f
            (ConvertTo-TgHtml -Text $(if ($Diff.OldDate) { $Diff.OldDate } else { 'не указана' })),
            (ConvertTo-TgHtml -Text $(if ($Article.Date) { $Article.Date } else { 'не указана' })))
    }

    if ($Diff.TimeChanged) {
        $Text += ''
        $Text += ('🕒 Время: <s>{0}</s> → {1}' -f
            (ConvertTo-TgHtml -Text $(if ($Diff.OldTime) { $Diff.OldTime } else { 'не указано' })),
            (ConvertTo-TgHtml -Text $(if ($Article.Time) { $Article.Time } else { 'не указано' })))
    }

    if (@($Diff.Added).Count -gt 0) {
        $Text += ''
        $Text += '➕ <b>Добавлено</b>'
        $Text += ('<blockquote>' + ((Format-ShortLines -Lines $Diff.Added) -join "`n") + '</blockquote>')
    }

    if (@($Diff.Removed).Count -gt 0) {
        $Text += ''
        $Text += '➖ <b>Убрано</b>'
        $Text += ('<blockquote>' + ((Format-ShortLines -Lines $Diff.Removed) -join "`n") + '</blockquote>')
    }

    return ($Text -join "`n")
}

# ============================================================
# ОТПРАВКА ПУБЛИКАЦИЙ И ЖУРНАЛ
# ============================================================

function Find-ArticleRecord {
    param(
        $State,
        [string]$Url
    )

    return (
        @($State.articles) |
        Where-Object { $_.url -eq $Url } |
        Select-Object -First 1
    )
}

function Send-NewArticle {
    param(
        $Article,
        $Config,
        $State
    )

    $Chats = [long[]]@($State.subscribers)

    if ($Chats.Count -eq 0) {
        Write-Warning ('Нет подписчиков Telegram. Не отправлена: {0}' -f $Article.Title)
        return $false
    }

    $Parts = @(Format-PostParts -Article $Article)
    $Buttons = Get-Buttons -Article $Article

    $Sent = Send-Parts `
        -Parts $Parts `
        -Buttons $Buttons `
        -Config $Config `
        -State $State `
        -ChatIds $Chats `
        -ReplyTo @{}

    if ($Sent.Count -eq 0) {
        return $false
    }

    $Pairs = @()

    foreach ($Key in $Sent.Keys) {
        $Pairs += [pscustomobject]@{
            chat = [long]$Key
            id = [long]$Sent[$Key]
        }
    }

    $Record = [pscustomobject]@{
        url = $Article.Url
        title = $Article.Title
        type = $Article.Type
        date = $Article.Date
        time = $Article.Time
        hash = $Article.Hash
        lines = @($Article.Lines)
        first_seen = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        updates = 0
        sent = @($Pairs)
    }

    $State.articles = @(
        @(@($State.articles) + @($Record)) |
        Select-Object -Last $MaxTrackedArticles
    )

    return $true
}

function Send-ChangeNotice {
    param(
        $Record,
        $Article,
        $Diff,
        $Config,
        $State
    )

    $Chats = [long[]]@($State.subscribers)

    if ($Chats.Count -eq 0) {
        return
    }

    # Отвечаем на прежнюю публикацию в каждом чате
    $ReplyTo = @{}

    foreach ($Pair in @($Record.sent)) {
        $ReplyTo[[string][long]$Pair.chat] = [long]$Pair.id
    }

    $Parts = @(Format-ChangeMessage -Article $Article -Diff $Diff)
    $Buttons = Get-Buttons -Article $Article -SiteOnly

    $null = Send-Parts `
        -Parts $Parts `
        -Buttons $Buttons `
        -Config $Config `
        -State $State `
        -ChatIds $Chats `
        -ReplyTo $ReplyTo
}

# ============================================================
# START
# ============================================================

Write-Host ''
Write-Host '========================================'
Write-Host ' Crimea Energy Telegram Monitor'
Write-Host '========================================'
Write-Host ''

$Config = Read-JsonFile -Path $ConfigPath -Default $null

if (-not $Config) {
    throw ('Не найден config.json: {0}' -f $ConfigPath)
}

if ([string]::IsNullOrWhiteSpace([string]$Config.telegram_bot_token)) {
    throw 'В config.json отсутствует telegram_bot_token.'
}

$StatePath = Join-Path $ScriptDirectory 'state.json'

$State = Read-JsonFile `
    -Path $StatePath `
    -Default (
        [pscustomobject]@{
            known_urls = @()
            sent_urls = @()
            initialized = $false
            subscribers = @()
            telegram_update_offset = 0
            articles = @()
        }
    )

$State = Ensure-StateProperties -State $State

Receive-Subscribers -Config $Config -State $State
Save-JsonFile -Object $State -Path $StatePath

Write-Host ('Telegram subscribers: {0}' -f @($State.subscribers).Count)

$NewsUrl = [string]$Config.news_url

if ([string]::IsNullOrWhiteSpace($NewsUrl)) {
    $NewsUrl = $DefaultNewsUrl
}

$NewsUri = [uri]$NewsUrl

Write-Host ('Reading news: {0}' -f $NewsUrl)

$NewsHtml = Get-PageWithRetry -Url $NewsUrl

if (-not $NewsHtml) {
    Write-Warning 'Страница новостей недоступна. Попробуем при следующем запуске.'
    Save-JsonFile -Object $State -Path $StatePath
    exit 0
}

$Links = @(Get-ArticleLinks -Html $NewsHtml -BaseUri $NewsUri)

Write-Host ('Article links found: {0}' -f $Links.Count)

if ($Links.Count -eq 0) {
    Write-Warning ('На странице {0} не найдено ссылок на статьи.' -f $NewsUrl)
    Save-JsonFile -Object $State -Path $StatePath
    exit 0
}

$KnownSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($Url in @($State.known_urls)) {
    if ($Url) { [void]$KnownSet.Add([string]$Url) }
}

$SentSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($Url in @($State.sent_urls)) {
    if ($Url) { [void]$SentSet.Add([string]$Url) }
}

function Save-AllState {
    $State.known_urls = @($KnownSet | Select-Object -Last 1000)
    $State.sent_urls = @($SentSet | Select-Object -Last 500)
    Save-JsonFile -Object $State -Path $StatePath
}

# ============================================================
# ПЕРВЫЙ ЗАПУСК
# ============================================================

if (-not [bool]$State.initialized) {

    Write-Host ''
    Write-Host 'FIRST RUN: ищем 3 последние подходящие публикации.'

    $Selected = @()

    foreach ($Link in ($Links | Select-Object -First 8)) {
        if ($Selected.Count -ge 3) { break }

        $Article = Get-Article -Url $Link.Url -FallbackTitle $Link.Title

        if (-not $Article) { continue }

        Write-Host ('Candidate: {0} [{1}]' -f $Article.Title, $Article.Type)

        if ($Article.Type -in $SendTypes) {
            $Selected += $Article
        }
    }

    # Было: новая -> старая. Отправляем: старая -> новая.
    [array]::Reverse($Selected)

    foreach ($Article in $Selected) {
        if (Send-NewArticle -Article $Article -Config $Config -State $State) {
            [void]$SentSet.Add($Article.Url)
            Write-Host ('Telegram: SENT - {0}' -f $Article.Title)
        }
    }

    # Всё, что было на сайте до первого запуска, считаем просмотренным
    foreach ($Link in $Links) {
        [void]$KnownSet.Add($Link.Url)
    }

    $State.initialized = $true
    Save-AllState

    Write-Host 'FIRST RUN COMPLETE.'
    exit 0
}

# ============================================================
# ОБЫЧНЫЙ ЗАПУСК
# ============================================================

Write-Host ''
Write-Host 'NORMAL RUN'

# --- 1. Новые публикации -------------------------------------

$NewLinks = @(
    $Links |
    Where-Object { -not $KnownSet.Contains($_.Url) }
)

Write-Host ('New URLs: {0}' -f $NewLinks.Count)

# На странице: новая -> старая. Обрабатываем: от старых новых к свежим.
[array]::Reverse($NewLinks)

foreach ($Link in $NewLinks) {

    Write-Host ''
    Write-Host ('NEW ARTICLE: {0}' -f $Link.Url)

    $Article = Get-Article -Url $Link.Url -FallbackTitle $Link.Title

    if (-not $Article) {
        # Страницы нет совсем - пропускаем навсегда. Иначе (сбой сайта) - попробуем в следующий запуск.
        if ($script:LastHttpStatus -in @(404, 410)) {
            [void]$KnownSet.Add($Link.Url)
            Save-AllState
            continue
        }

        Write-Warning 'Статья не загрузилась, остановка. Повтор при следующем запуске.'
        break
    }

    Write-Host ('Title: {0}' -f $Article.Title)
    Write-Host ('Type: {0}' -f $Article.Type)

    if ($Article.Type -in $SendTypes) {
        if (Send-NewArticle -Article $Article -Config $Config -State $State) {
            Write-Host 'Telegram: SENT'
            [void]$SentSet.Add($Article.Url)
        }
    }
    else {
        Write-Host 'Telegram: NOT SENT (тип отключён в настройках)'
    }

    [void]$KnownSet.Add($Link.Url)
    Save-AllState
}

# --- 2. Изменения в уже отправленных публикациях -------------

$Now = [DateTimeOffset]::UtcNow
$NowSeconds = $Now.ToUnixTimeSeconds()
$MinuteNow = $Now.Minute

$Watch = @()

foreach ($Record in (@($State.articles) | Sort-Object { [long]$_.first_seen } -Descending)) {
    $AgeHours = ($NowSeconds - [long]$Record.first_seen) / 3600.0

    if ($AgeHours -gt $WatchHours) { continue }

    # Свежие проверяем каждый запуск, более старые - примерно раз в час
    if (($AgeHours -gt $FastWatchHours) -and ($MinuteNow -ge 10)) { continue }

    $Watch += $Record
}

$Watch = @($Watch | Select-Object -First $MaxWatchedPerRun)

Write-Host ''
Write-Host ('Проверка изменений, публикаций: {0}' -f $Watch.Count)

foreach ($Record in $Watch) {

    $Article = Get-Article -Url $Record.url -FallbackTitle $Record.title

    # Не загрузилась или пустая страница - ничего не трогаем
    if (-not $Article) { continue }
    if (@($Article.Lines).Count -eq 0) { continue }

    if ($Article.Hash -eq $Record.hash) { continue }

    $Diff = Compare-Article -Record $Record -Article $Article

    if ($Diff.HasChanges -and ([int]$Record.updates -lt $MaxUpdatesPerArticle)) {
        Write-Host ('CHANGED: {0}' -f $Article.Title)

        Send-ChangeNotice `
            -Record $Record `
            -Article $Article `
            -Diff $Diff `
            -Config $Config `
            -State $State

        $Record.updates = [int]$Record.updates + 1
    }

    # Запоминаем новую версию
    $Record.title = $Article.Title
    $Record.type = $Article.Type
    $Record.date = $Article.Date
    $Record.time = $Article.Time
    $Record.hash = $Article.Hash
    $Record.lines = @($Article.Lines)

    Save-AllState
}

Save-AllState

Write-Host ''
Write-Host 'NORMAL RUN COMPLETE.'
