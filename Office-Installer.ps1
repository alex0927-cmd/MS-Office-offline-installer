# Office 2021 Professional Plus — об'єднаний інсталятор
# Замінює: Init-ConsoleEncoding, Resolve-OfficePaths, Prepare-OfficeConfig,
#         Prepare-InstallConfig, Download-OfficePackage, Install-Office2021,
#         Verify-OfficeInstall, Verify-OfficeRemoved

param(
    [ValidateSet('PrepareConfig', 'PrepareInstallConfig', 'Download', 'DownloadOdt', 'Install', 'VerifyInstall', 'VerifyRemoved', 'ResolvePaths', 'GetInstalledOffice', 'CheckTeamsSkype', 'RemoveTeamsSkype')]
    [string]$Command = 'Install',

    [string]$InstallDir,
    [string]$ScriptDir,
    [string]$BasePath,

    [ValidateSet('Install', 'Reinstall', 'Uninstall', 'Download', 'RemoveApps')]
    [string]$Mode,

    [ValidateSet('Full', 'Compact', 'Brief')]
    [string]$OutputFormat = 'Full',

    [switch]$Reinstall,

    [int]$SetupExitCode = 0,
    [switch]$AlreadyInstalled
)

# --- Init-ConsoleEncoding 

function Initialize-UkrainianConsole {
    cmd /c "chcp 1251 >nul 2>&1"
    $enc = [System.Text.Encoding]::GetEncoding(1251)
    [Console]::OutputEncoding = $enc
    [Console]::InputEncoding = $enc
    $global:OutputEncoding = $enc
}

# --- Resolve-OfficePaths 

function Resolve-OfficePaths {
    param(
        [string]$BasePath = $PSScriptRoot
    )

    $scriptDir = (Resolve-Path -LiteralPath $BasePath).Path

    [PSCustomObject]@{
        ScriptDir    = $scriptDir
        RootDir      = $scriptDir
        SetupPresent = Test-Path -LiteralPath (Join-Path $scriptDir 'setup.exe')
    }
}

# --- configurations.xml helpers 

function Get-ConfigurationsFile {
    param([string]$Dir)

    foreach ($searchDir in @($Dir, $PSScriptRoot)) {
        if (-not $searchDir) { continue }

        $combined = Join-Path $searchDir 'configurations.xml'
        if (Test-Path -LiteralPath $combined) { return $combined }

        $legacy = Join-Path $searchDir 'configuration.xml'
        if (Test-Path -LiteralPath $legacy) { return $legacy }
    }

    Write-Error "Не знайдено configurations.xml поруч зі скриптами ($PSScriptRoot)"
    exit 1
}

function Get-OdtDownloadUrl {
    param([string]$Dir)

    $defaultUrl = 'https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_20026-20112.exe'

    foreach ($searchDir in @($Dir, $PSScriptRoot)) {
        if (-not $searchDir) { continue }
        $configFile = Join-Path $searchDir 'configurations.xml'
        if (-not (Test-Path -LiteralPath $configFile)) { continue }

        [xml]$doc = Get-Content -LiteralPath $configFile
        if ($doc.OfficeConfigurations.OdtDownload.Url) {
            return [string]$doc.OfficeConfigurations.OdtDownload.Url
        }
    }

    return $defaultUrl
}

function Get-ExpectedDownloadBytes {
    param(
        [string]$Dir,
        [string]$OfficePath
    )

    $defaultBytes = 2200000000L

    foreach ($searchDir in @($Dir, $PSScriptRoot)) {
        if (-not $searchDir) { continue }
        $configFile = Join-Path $searchDir 'configurations.xml'
        if (-not (Test-Path -LiteralPath $configFile)) { continue }

        [xml]$doc = Get-Content -LiteralPath $configFile
        if ($doc.OfficeConfigurations.DownloadProfile.ExpectedBytes) {
            $defaultBytes = [long]$doc.OfficeConfigurations.DownloadProfile.ExpectedBytes
            break
        }
    }

    if ($OfficePath -and (Test-Path -LiteralPath $OfficePath)) {
        $existing = Get-PackageSize -Path $OfficePath
        if ($existing -gt 100MB) {
            $defaultBytes = [Math]::Max($defaultBytes, [long]($existing * 1.02))
        }
    }

    return $defaultBytes
}

function Get-BaseConfiguration {
    param([string]$InstallDir)

    $configFile = Get-ConfigurationsFile -Dir $InstallDir
    [xml]$doc = Get-Content -LiteralPath $configFile

    if ($doc.DocumentElement.LocalName -eq 'OfficeConfigurations') {
        $node = $doc.OfficeConfigurations.Configuration | Where-Object { $_.id -eq 'configuration' } | Select-Object -First 1
        if (-not $node) {
            Write-Error "У configurations.xml відсутній блок id='configuration'"
            exit 1
        }
        return $node
    }

    return $doc.Configuration
}

# --- Prepare-OfficeConfig 

function Add-OdtProductNode {
    param(
        [System.Xml.XmlDocument]$Doc,
        [System.Xml.XmlElement]$AddElement,
        $SrcProduct
    )

    $product = $Doc.CreateElement('Product')
    [void]$product.SetAttribute('ID', $SrcProduct.ID)

    foreach ($lang in @($SrcProduct.Language)) {
        if (-not $lang) { continue }
        $langNode = $Doc.CreateElement('Language')
        [void]$langNode.SetAttribute('ID', $lang.ID)
        [void]$product.AppendChild($langNode)
    }

    foreach ($exclude in @($SrcProduct.ExcludeApp)) {
        if (-not $exclude) { continue }
        $exNode = $Doc.CreateElement('ExcludeApp')
        [void]$exNode.SetAttribute('ID', $exclude.ID)
        [void]$product.AppendChild($exNode)
    }

    [void]$AddElement.AppendChild($product)
}

function Invoke-PrepareOfficeConfig {
    param(
        [string]$InstallDir,
        [ValidateSet('Install', 'Reinstall', 'Uninstall', 'Download', 'RemoveApps')]
        [string]$Mode
    )

    if (-not $InstallDir) {
        $paths = Resolve-OfficePaths -BasePath $PSScriptRoot
        $InstallDir = $paths.RootDir
    }

    $src = Get-BaseConfiguration -InstallDir $InstallDir

    $doc = New-Object System.Xml.XmlDocument
    [void]$doc.LoadXml('<Configuration></Configuration>')
    $cfg = $doc.DocumentElement

    switch ($Mode) {
        'Download' {
            $add = $doc.CreateElement('Add')
            [void]$add.SetAttribute('OfficeClientEdition', $src.Add.OfficeClientEdition)
            [void]$add.SetAttribute('Channel', $src.Add.Channel)

            Add-OdtProductNode -Doc $doc -AddElement $add -SrcProduct $src.Add.Product
            [void]$cfg.AppendChild($add)

            $display = $doc.CreateElement('Display')
            [void]$display.SetAttribute('Level', 'None')
            [void]$display.SetAttribute('AcceptEULA', 'TRUE')
            [void]$cfg.AppendChild($display)

            $fileName = 'configuration-download.xml'
        }
        'Uninstall' {
            $remove = $doc.CreateElement('Remove')
            [void]$remove.SetAttribute('All', 'TRUE')
            [void]$cfg.AppendChild($remove)

            $display = $doc.CreateElement('Display')
            [void]$display.SetAttribute('Level', 'Full')
            [void]$display.SetAttribute('AcceptEULA', 'TRUE')
            [void]$cfg.AppendChild($display)

            $fileName = 'configuration-remove.xml'
        }
        'RemoveApps' {
            $add = $doc.CreateElement('Add')
            [void]$add.SetAttribute('OfficeClientEdition', $src.Add.OfficeClientEdition)
            [void]$add.SetAttribute('Channel', $src.Add.Channel)

            $product = $doc.CreateElement('Product')
            [void]$product.SetAttribute('ID', $src.Add.Product.ID)

            foreach ($lang in @($src.Add.Product.Language)) {
                if (-not $lang) { continue }
                $langNode = $doc.CreateElement('Language')
                [void]$langNode.SetAttribute('ID', $lang.ID)
                [void]$product.AppendChild($langNode)
            }

            foreach ($appId in @('Lync', 'Teams')) {
                $exNode = $doc.CreateElement('ExcludeApp')
                [void]$exNode.SetAttribute('ID', $appId)
                [void]$product.AppendChild($exNode)
            }

            [void]$add.AppendChild($product)
            [void]$cfg.AppendChild($add)

            $display = $doc.CreateElement('Display')
            [void]$display.SetAttribute('Level', 'None')
            [void]$display.SetAttribute('AcceptEULA', 'TRUE')
            [void]$cfg.AppendChild($display)

            $fileName = 'configuration-remove-teams-skype.xml'
        }
        default {
            if ($Mode -eq 'Reinstall') {
                $remove = $doc.CreateElement('Remove')
                [void]$remove.SetAttribute('All', 'TRUE')
                [void]$cfg.AppendChild($remove)
            }

            $add = $doc.CreateElement('Add')
            [void]$add.SetAttribute('OfficeClientEdition', $src.Add.OfficeClientEdition)
            [void]$add.SetAttribute('Channel', $src.Add.Channel)
            [void]$add.SetAttribute('SourcePath', $InstallDir)

            Add-OdtProductNode -Doc $doc -AddElement $add -SrcProduct $src.Add.Product
            [void]$cfg.AppendChild($add)

            $display = $doc.CreateElement('Display')
            [void]$display.SetAttribute('Level', 'Full')
            [void]$display.SetAttribute('AcceptEULA', 'TRUE')
            [void]$cfg.AppendChild($display)

            $fileName = 'configuration-install.xml'
        }
    }

    $outPath = Join-Path $InstallDir $fileName
    $doc.Save($outPath)
    Write-Output $outPath
}

# --- Prepare-InstallConfig (legacy) 

function Invoke-PrepareInstallConfig {
    param(
        [string]$InstallDir,
        [switch]$Reinstall
    )

    if (-not $InstallDir) {
        $paths = Resolve-OfficePaths -BasePath $PSScriptRoot
        $InstallDir = $paths.RootDir
    }

    $configFile = Get-ConfigurationsFile -Dir $InstallDir
    [xml]$doc = Get-Content -LiteralPath $configFile

    if ($doc.DocumentElement.LocalName -eq 'OfficeConfigurations') {
        $baseNode = $doc.OfficeConfigurations.Configuration | Where-Object { $_.id -eq 'configuration' } | Select-Object -First 1
        $inner = $baseNode.OuterXml
        [void]$doc.LoadXml("<Configuration>$inner</Configuration>")
    }

    $cfg = $doc.Configuration

    $cfg.Add.SourcePath = $InstallDir
    $cfg.Display.Level = 'Full'
    $cfg.Display.AcceptEULA = 'TRUE'

    if ($Reinstall) {
        $remove = $doc.CreateElement('Remove')
        $remove.SetAttribute('All', 'TRUE')
        if ($cfg.Remove) {
            $cfg.Remove.ParentNode.ReplaceChild($remove, $cfg.Remove) | Out-Null
        } else {
            [void]$cfg.InsertBefore($remove, $cfg.Add)
        }
    }

    $outPath = Join-Path $InstallDir 'configuration-install.xml'
    $doc.Save($outPath)
    Write-Output $outPath
}

# --- Download-Odt (setup.exe) 

function Invoke-DownloadOdt {
    param(
        [string]$TargetDir
    )

    Initialize-UkrainianConsole
    $ErrorActionPreference = 'Stop'

    if (-not $TargetDir) {
        $paths = Resolve-OfficePaths -BasePath $PSScriptRoot
        $TargetDir = $paths.RootDir
    }

    $setupPath = Join-Path $TargetDir 'setup.exe'

    Write-Host ""
    Write-Host "  [i] Перевірка інтернет-з'єднання з сервером Microsoft..." -ForegroundColor Cyan

    $online = $false
    try {
        $online = Test-Connection -ComputerName 'www.microsoft.com' -Count 1 -Quiet -ErrorAction Stop
    } catch {
        $online = $false
    }

    if (-not $online) {
        Write-Host "  [X] Інтернет недоступний. Завантаження неможливе." -ForegroundColor Red
        exit 1
    }
    Write-Host "  [OK] Інтернет доступний" -ForegroundColor Green

    if (Test-Path -LiteralPath $setupPath) {
        Write-Host ""
        Write-Host "  [i] setup.exe уже існує у $TargetDir" -ForegroundColor Yellow
    }

    $url = Get-OdtDownloadUrl -Dir $TargetDir
    $tempExe = Join-Path $env:TEMP ("odt_{0}.exe" -f ([guid]::NewGuid().ToString('N').Substring(0, 8)))

    Write-Host ""
    Write-Host "  Завантаження Office Deployment Tool..." -ForegroundColor Cyan
    Write-Host "  Джерело: Microsoft" -ForegroundColor Gray
    Write-Host "  Куди:    $TargetDir" -ForegroundColor Gray
    Write-Host ""

    try {
        Invoke-WebRequest -Uri $url -OutFile $tempExe -UseBasicParsing
    } catch {
        Write-Host "  [X] Не вдалося завантажити ODT: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }

    Write-Host "  Розпакування setup.exe..." -ForegroundColor Cyan

    $extractArgs = "/extract:`"$TargetDir`"", '/quiet'
    $proc = Start-Process -FilePath $tempExe -ArgumentList $extractArgs -Wait -PassThru -NoNewWindow

    Remove-Item -LiteralPath $tempExe -Force -ErrorAction SilentlyContinue

    Write-Host ""

    if ((Test-Path -LiteralPath $setupPath) -and $proc.ExitCode -eq 0) {
        Write-Host "  [OK] setup.exe готовий до роботи" -ForegroundColor Green
        Write-Host "       $setupPath" -ForegroundColor Green
        exit 0
    }

    Write-Host "  [X] setup.exe не з'явився (код розпакування: $($proc.ExitCode))" -ForegroundColor Red
    exit 1
}

# --- Download-OfficePackage 

function Format-Size {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N0} MB' -f ($Bytes / 1MB)) }
    return ('{0:N0} KB' -f ($Bytes / 1KB))
}

function Get-PackageSize {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return 0 }
    return (Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
        Measure-Object -Property Length -Sum).Sum
}

function Get-PackageFileCount {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return 0 }
    return @(Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue).Count
}

function Format-TransferRate {
    param([double]$BytesPerSec)
    if ($BytesPerSec -ge 1GB) { return ('{0:N2} GB/s' -f ($BytesPerSec / 1GB)) }
    if ($BytesPerSec -ge 1MB) { return ('{0:N1} MB/s' -f ($BytesPerSec / 1MB)) }
    if ($BytesPerSec -ge 1KB) { return ('{0:N0} KB/s' -f ($BytesPerSec / 1KB)) }
    if ($BytesPerSec -gt 0) { return ('{0:N0} B/s' -f $BytesPerSec) }
    return '0 B/s'
}

function Show-PackageDownloadProgress {
    param(
        [long]$Downloaded,
        [long]$Total,
        [double]$Speed,
        [TimeSpan]$Elapsed,
        [int]$FileCount,
        [switch]$Final
    )

    $pct = 0.0
    if ($Total -gt 0) {
        $pct = if ($Final) { 100.0 } else { [Math]::Min(99.9, ($Downloaded / $Total) * 100.0) }
    }

    $barLen = 30
    $fill = [int][Math]::Round($barLen * $pct / 100.0)
    if ($fill -gt $barLen) { $fill = $barLen }
    $bar = ('#' * $fill) + ('-' * ($barLen - $fill))

    $etaStr = ''
    if (-not $Final -and $Speed -gt 0 -and $Total -gt $Downloaded) {
        $etaSec = ($Total - $Downloaded) / $Speed
        if ($etaSec -lt 86400) {
            $etaStr = '  ~' + ([TimeSpan]::FromSeconds($etaSec).ToString('hh\:mm\:ss'))
        }
    }

    $line = ('  [{0}] {1,5:N1}%  {2} / {3}  {4}  {5}{6}  {7} files' -f `
        $bar, $pct, `
        (Format-Size $Downloaded), `
        (Format-Size $Total), `
        (Format-TransferRate $Speed), `
        ('{0:hh\:mm\:ss}' -f $Elapsed), `
        $etaStr, `
        $FileCount)

    if ($Final) {
        Write-Host $line -ForegroundColor Green
    } else {
        Write-Host ("`r{0,-120}" -f $line) -NoNewline
    }
}

function Invoke-DownloadOfficePackage {
    param(
        [string]$InstallDir,
        [string]$ScriptDir
    )

    Initialize-UkrainianConsole
    $ErrorActionPreference = 'Stop'

    if (-not $ScriptDir) { $ScriptDir = $PSScriptRoot }
    if (-not $InstallDir) {
        $paths = Resolve-OfficePaths -BasePath $ScriptDir
        $InstallDir = $paths.RootDir
        $ScriptDir = $paths.ScriptDir
    }

    if (-not (Test-Path -LiteralPath (Join-Path $InstallDir 'setup.exe'))) {
        Write-Host "  [X] setup.exe не знайдено у $InstallDir" -ForegroundColor Red
        Write-Host "      Спочатку виконайте пункт [1] — завантажити ODT (setup.exe)" -ForegroundColor Yellow
        exit 1
    }

    Write-Host ""
    Write-Host "  [i] Перевірка інтернет-з'єднання..." -ForegroundColor Cyan

    $online = $false
    try {
        $online = Test-Connection -ComputerName 'www.microsoft.com' -Count 1 -Quiet -ErrorAction Stop
    } catch {
        $online = $false
    }

    if (-not $online) {
        Write-Host "  [X] Інтернет недоступний. Завантаження неможливе." -ForegroundColor Red
        exit 1
    }
    Write-Host "  [OK] Інтернет доступний" -ForegroundColor Green

    $setup = Join-Path $InstallDir 'setup.exe'
    $config = Join-Path $InstallDir 'configuration-download.xml'
    $dataPath = Join-Path $InstallDir 'Office\Data'

    Invoke-PrepareOfficeConfig -InstallDir $InstallDir -Mode Download | Out-Null

    Write-Host ""
    Write-Host "  Продукт:  Office LTSC Professional Plus 2021" -ForegroundColor Yellow
    Write-Host "  Мова:     uk-ua" -ForegroundColor Yellow
    Write-Host "  Без:      Skype, Microsoft Teams" -ForegroundColor Yellow
    Write-Host "  Куди:     $dataPath" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Завантаження розпочато. Це може зайняти 10-30 хвилин." -ForegroundColor Cyan
    Write-Host ""

    $officePath = Join-Path $InstallDir 'Office'
    $expectedBytes = Get-ExpectedDownloadBytes -Dir $InstallDir -OfficePath $officePath
    Write-Host ("  Очікуваний розмір: ~{0}" -f (Format-Size $expectedBytes)) -ForegroundColor Gray
    Write-Host ""

    $startTime = Get-Date

    $job = Start-Job -ScriptBlock {
        param($SetupPath, $ConfigPath, $WorkDir)
        Set-Location $WorkDir
        & $SetupPath /download $ConfigPath
        return $LASTEXITCODE
    } -ArgumentList $setup, $config, $InstallDir

    $lastSize = 0L
    $lastTick = Get-Date
    $stableSize = 0L
    $stableSince = Get-Date

    while ($job.State -eq 'Running') {
        Start-Sleep -Seconds 2

        $currentSize = [long](Get-PackageSize -Path $officePath)
        $fileCount = Get-PackageFileCount -Path $officePath
        $now = Get-Date
        $elapsed = $now - $startTime
        $tickSec = ($now - $lastTick).TotalSeconds

        if ($currentSize -gt $expectedBytes) {
            $expectedBytes = [long]($currentSize * 1.02)
        }

        $speed = 0.0
        if ($tickSec -gt 0) {
            $speed = ($currentSize - $lastSize) / $tickSec
        }

        if ($currentSize -eq $stableSize) {
            if (($now - $stableSince).TotalSeconds -ge 30 -and $currentSize -gt 50MB) {
                $expectedBytes = [Math]::Max($expectedBytes, $currentSize)
            }
        } else {
            $stableSize = $currentSize
            $stableSince = $now
        }

        Show-PackageDownloadProgress -Downloaded $currentSize -Total $expectedBytes `
            -Speed $speed -Elapsed $elapsed -FileCount $fileCount

        $lastSize = $currentSize
        $lastTick = $now
    }

    $exitCode = Receive-Job $job
    Remove-Job $job -Force

    $finalSize = [long](Get-PackageSize -Path $officePath)
    $fileCount = Get-PackageFileCount -Path $officePath
    $elapsedTotal = (Get-Date) - $startTime
    $avgSpeed = if ($elapsedTotal.TotalSeconds -gt 0) { $finalSize / $elapsedTotal.TotalSeconds } else { 0.0 }

    Show-PackageDownloadProgress -Downloaded $finalSize -Total $finalSize `
        -Speed $avgSpeed -Elapsed $elapsedTotal -FileCount $fileCount -Final

    Write-Host ""

    if ($exitCode -eq 0 -and $finalSize -gt 100MB) {
        Write-Host "  [OK] Завантаження завершено успішно" -ForegroundColor Green
        Write-Host "       Час:     $($elapsedTotal.ToString('hh\:mm\:ss'))" -ForegroundColor Green
        Write-Host "       Розмір:  $(Format-Size $finalSize)" -ForegroundColor Green
        Write-Host "       Файлів:  $fileCount" -ForegroundColor Green
        Write-Host ""
        Write-Host "  Скопіюйте папку пакету:" -ForegroundColor Yellow
        Write-Host "    $InstallDir" -ForegroundColor Yellow
        Write-Host "  на USB або інший ПК без інтернету." -ForegroundColor Yellow
        exit 0
    }

    Write-Host "  [X] Завантаження не вдалося (код $exitCode)" -ForegroundColor Red
    if ($finalSize -gt 0) {
        Write-Host "      Завантажено частково: $(Format-Size $finalSize)" -ForegroundColor Yellow
    }
    Write-Host "      Перевірте логи у %TEMP%" -ForegroundColor Yellow
    exit 1
}

# --- Install-Office2021 

function Invoke-InstallOffice2021 {
    Initialize-UkrainianConsole
    $ErrorActionPreference = 'Stop'
    $paths = Resolve-OfficePaths -BasePath $PSScriptRoot
    $root = $paths.RootDir

    if (-not $paths.SetupPresent) {
        Write-Host "setup.exe не знайдено. Спочатку виконайте пункт [1] у меню." -ForegroundColor Red
        exit 1
    }

    Set-Location $root

    Invoke-PrepareOfficeConfig -InstallDir $root -Mode Install | Out-Null

    Write-Host "Встановлення Office 2021 Pro Plus..." -ForegroundColor Cyan
    Write-Host "Пакет: $root" -ForegroundColor Gray
    & (Join-Path $root 'setup.exe') /configure (Join-Path $root 'configuration-install.xml')

    if ($LASTEXITCODE -eq 0) {
        Write-Host "`nOffice 2021 Pro Plus встановлено успішно!" -ForegroundColor Green
    } else {
        Write-Host "`nПомилка встановлення. Код: $LASTEXITCODE" -ForegroundColor Red
        Write-Host "Перевірте лог: $env:TEMP" -ForegroundColor Yellow
    }
}

# --- Verify-OfficeInstall / Verify-OfficeRemoved 

function Get-OfficeEditionLabel {
    param(
        [string]$ProductId,
        [string]$DisplayName
    )

    $map = @{
        'ProPlus2021Volume'      = 'Office LTSC Professional Plus 2021'
        'ProPlus2024Volume'      = 'Office LTSC Professional Plus 2024'
        'ProPlus2019Volume'      = 'Office Professional Plus 2019'
        'ProPlus2016Volume'      = 'Office Professional Plus 2016'
        'Standard2021Volume'     = 'Office LTSC Standard 2021'
        'Standard2024Volume'     = 'Office LTSC Standard 2024'
        'O365ProPlusRetail'      = 'Microsoft 365 Apps for enterprise'
        'O365BusinessRetail'     = 'Microsoft 365 Apps for business'
        'O365HomePremRetail'     = 'Microsoft 365 для сім''ї'
        'ProjectPro2021Volume'   = 'Project Professional 2021'
        'VisioPro2021Volume'     = 'Visio Professional 2021'
    }

    if ($ProductId) {
        foreach ($key in $map.Keys) {
            if ($ProductId -match [regex]::Escape($key)) { return $map[$key] }
        }
    }

    if ($DisplayName -match '2024') { return 'Office LTSC 2024' }
    if ($DisplayName -match '2021') { return 'Office LTSC 2021' }
    if ($DisplayName -match '2019') { return 'Office 2019' }
    if ($DisplayName -match '2016') { return 'Office 2016' }
    if ($DisplayName -match '365')  { return 'Microsoft 365' }

    return $DisplayName
}

function Get-OfficeChannelLabel {
    param(
        $ClickToRun,
        [string]$ProductId
    )

    $channel = if ($ClickToRun) { [string]$ClickToRun.UpdateChannel } else { '' }

    if ($channel -match '[\\/]' -or $channel -match '^[A-Za-z]:') {
        $channel = ''
    }

    if ($channel) { return $channel }

    if ($ProductId -match '2024Volume') { return 'PerpetualVL2024' }
    if ($ProductId -match '2021Volume') { return 'PerpetualVL2021' }
    if ($ProductId -match '2019Volume') { return 'PerpetualVL2019' }
    if ($ProductId -match '2016Volume') { return 'PerpetualVL2016' }
    if ($ProductId -match 'Retail')       { return 'Current (Microsoft 365)' }

    return ''
}

function Get-InstalledOfficeInfo {
    $uninstallKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $entries = @(Get-ItemProperty $uninstallKeys -ErrorAction SilentlyContinue |
        Where-Object { $_ -and $_.DisplayName } |
        Where-Object {
            $_.DisplayName -match 'Microsoft Office|Office LTSC|Microsoft 365|Office Professional Plus|Office Home|Project Professional|Visio Professional' -and
            $_.DisplayName -notmatch 'Component|Extensibility|Licensing|Proofing|Language Interface|Update for|Filter Pack|Groove|Lync'
        })

    $seen = @{}
    $entries = $entries | Where-Object {
        if ($seen.ContainsKey($_.DisplayName)) { return $false }
        $seen[$_.DisplayName] = $true
        return $true
    }

    $ctr = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
    $productId = if ($ctr) { [string]$ctr.ProductReleaseIds } else { '' }
    $edition = Get-OfficeEditionLabel -ProductId $productId -DisplayName ''

    $results = foreach ($entry in $entries) {
        $version = [string]$entry.DisplayVersion
        if ($ctr -and $ctr.VersionToReport) {
            $version = [string]$ctr.VersionToReport
        }

        $build = if ($ctr -and $ctr.ClientVersionToReport) { [string]$ctr.ClientVersionToReport } else { $version }
        $platform = if ($ctr -and $ctr.Platform) { [string]$ctr.Platform } else { '' }
        $channel = Get-OfficeChannelLabel -ClickToRun $ctr -ProductId $productId

        [PSCustomObject]@{
            DisplayName       = [string]$entry.DisplayName
            DisplayVersion    = $version
            BuildVersion      = $build
            InstallLocation   = [string]$entry.InstallLocation
            ProductReleaseIds = $productId
            Edition           = Get-OfficeEditionLabel -ProductId $productId -DisplayName $entry.DisplayName
            Channel           = $channel
            Platform          = $platform
        }
    }

    return @($results | Sort-Object {
        if ($_.DisplayName -match 'Professional Plus|Microsoft 365 Apps|Office LTSC') { 0 }
        elseif ($_.DisplayName -match 'Microsoft Office') { 1 }
        else { 2 }
    }, @{ Expression = 'DisplayName' })
}

function Invoke-GetInstalledOffice {
    param(
        [ValidateSet('Full', 'Compact', 'Brief')]
        [string]$Format = 'Full'
    )

    if ($Format -ne 'Brief') { Initialize-UkrainianConsole }
    $items = @(Get-InstalledOfficeInfo)

    if ($items.Count -eq 0) {
        switch ($Format) {
            'Brief'   { return }
            'Compact' { Write-Host '  Office: [не встановлено]' -ForegroundColor Gray }
            'Full'    { Write-StatusLine 'i' 'Office у системі не знайдено' 'Yellow' }
        }
        return
    }

    switch ($Format) {
        'Brief' {
            $i = $items[0]
            $name = ($i.DisplayName -replace '\|', '/')
            Write-Output ('{0}|{1}|{2}|{3}|{4}' -f $name, $i.DisplayVersion, $i.Platform, $i.Channel, $i.Edition)
        }
        'Compact' {
            foreach ($i in $items) {
                $extra = @($i.Platform, $(if ($i.Channel) { $i.Channel } else { $i.Edition })) -ne '' -join ', '
                Write-Host ("  Office: {0}  v{1}  ({2})" -f $i.DisplayName, $i.DisplayVersion, $extra) -ForegroundColor Cyan
            }
        }
        'Full' {
            foreach ($i in $items) {
                Write-Host ""
                Write-StatusLine 'OK' 'Office знайдено в системі' 'Green'
                Write-Host "    Назва:     $($i.DisplayName)"
                Write-Host "    Видання:   $($i.Edition)"
                Write-Host "    Версія:    $($i.DisplayVersion)"
                if ($i.BuildVersion -and $i.BuildVersion -ne $i.DisplayVersion) {
                    Write-Host "    Білд:      $($i.BuildVersion)"
                }
                if ($i.Platform)  { Write-Host "    Платформа: $($i.Platform)" }
                if ($i.Channel)   { Write-Host "    Канал:     $($i.Channel)" }
                if ($i.ProductReleaseIds) { Write-Host "    Продукт:   $($i.ProductReleaseIds)" }
                if ($i.InstallLocation)   { Write-Host "    Шлях:      $($i.InstallLocation)" }
            }
        }
    }
}

function Write-StatusLine {
    param([string]$Symbol, [string]$Text, [string]$Color = 'White')
    Write-Host "  [$Symbol] " -NoNewline -ForegroundColor $Color
    Write-Host $Text
}

function Invoke-VerifyOfficeInstall {
    param(
        [int]$SetupExitCode = 0,
        [switch]$AlreadyInstalled
    )

    Initialize-UkrainianConsole
    $ErrorActionPreference = 'SilentlyContinue'

    Write-Host ""
    Write-StatusLine "i" "Перевірка реєстру Windows..." "Cyan"

    $officeItems = @(Get-InstalledOfficeInfo)
    $officeEntry = $officeItems | Select-Object -First 1

    $officeRoots = @(
        'C:\Program Files\Microsoft Office\root\Office16',
        'C:\Program Files\Microsoft Office\Office16'
    )

    $appMap = @{
        'Word' = 'WINWORD.EXE'
        'Excel' = 'EXCEL.EXE'
        'PowerPoint' = 'POWERPNT.EXE'
        'Outlook' = 'OUTLOOK.EXE'
        'Access' = 'MSACCESS.EXE'
        'Publisher' = 'MSPUB.EXE'
        'OneNote' = 'ONENOTE.EXE'
    }

    $installedApps = @()
    foreach ($root in $officeRoots) {
        foreach ($name in $appMap.Keys) {
            if (Test-Path (Join-Path $root $appMap[$name])) {
                if ($installedApps -notcontains $name) { $installedApps += $name }
            }
        }
    }

    Write-Host ""
    if ($officeItems.Count -gt 0) {
        foreach ($item in $officeItems) {
            Write-StatusLine "OK" "Office знайдено в системі" "Green"
            Write-Host ""
            Write-Host "  Що встановлено:" -ForegroundColor Yellow
            Write-Host "    Назва:     $($item.DisplayName)"
            Write-Host "    Видання:   $($item.Edition)"
            Write-Host "    Версія:    $($item.DisplayVersion)"
            if ($item.BuildVersion -and $item.BuildVersion -ne $item.DisplayVersion) {
                Write-Host "    Білд:      $($item.BuildVersion)"
            }
            if ($item.Platform) { Write-Host "    Платформа: $($item.Platform)" }
            if ($item.Channel)  { Write-Host "    Канал:     $($item.Channel)" }
            if ($item.ProductReleaseIds) { Write-Host "    Продукт:   $($item.ProductReleaseIds)" }
            Write-Host "    Шлях:      $($item.InstallLocation)"
            Write-Host ""
        }
        if ($installedApps.Count -gt 0) {
            Write-StatusLine "OK" "Програми: $($installedApps -join ', ')" "Green"
        } else {
            Write-StatusLine "?" "Програми Office не знайдено" "Yellow"
        }
    } else {
        Write-StatusLine "X" "Office НЕ знайдено в системі" "Red"
        if ($SetupExitCode -eq 0) {
            Write-Host ""
            Write-Host "  setup.exe повернув код 0, але Office у реєстрі відсутній." -ForegroundColor Yellow
            Write-Host "  Можливо, встановлення скасовано у вікні інсталятора." -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "  Статус активації:" -ForegroundColor Yellow

    $ospp = 'C:\Program Files\Microsoft Office\Office16\OSPP.VBS'
    if (-not (Test-Path $ospp)) {
        foreach ($root in $officeRoots) {
            $candidate = Join-Path $root 'OSPP.VBS'
            if (Test-Path $candidate) { $ospp = $candidate; break }
        }
    }

    if (Test-Path $ospp) {
        $job = Start-Job -ScriptBlock {
            param($Path)
            & cscript.exe //nologo $Path /dstatus 2>&1 | Out-String
        } -ArgumentList $ospp

        $completed = Wait-Job $job -Timeout 15
        if ($completed) {
            $statusRaw = Receive-Job $job
            Remove-Job $job -Force

            if ($statusRaw -match 'LICENSE STATUS:\s+([^\r\n]+)') {
                $licenseStatus = $Matches[1].Trim()
                if ($licenseStatus -eq '---LICENSED---') {
                    Write-StatusLine "OK" "Office активовано" "Green"
                } elseif ($licenseStatus -eq '---UNLICENSED---') {
                    Write-StatusLine "!" "Office НЕ активовано — введіть ключ продукту" "Yellow"
                } elseif ($licenseStatus -eq '---NOTIFICATIONS---') {
                    Write-StatusLine "!" "Office потребує активації — введіть ключ продукту" "Yellow"
                    if ($statusRaw -match 'Last 5 characters of installed product key:\s+(\S+)') {
                        Write-Host "    (останні 5 символів ключа: $($Matches[1]))"
                    }
                } else {
                    Write-StatusLine "?" "Статус ліцензії: $licenseStatus" "Yellow"
                }
            } else {
                Write-StatusLine "?" "Не вдалося визначити статус ліцензії" "Yellow"
            }
        } else {
            Stop-Job $job -Force
            Remove-Job $job -Force
            Write-StatusLine "?" "Перевірка активації занадто довга — перевірте в Word/Excel" "Yellow"
        }
    } else {
        Write-StatusLine "?" "OSPP.VBS не знайдено" "Yellow"
    }

    Write-Host ""
    Write-Host "  Ярлики в Меню Пуск:" -ForegroundColor Yellow
    $startMenu = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
    $shortcuts = Get-ChildItem -Path $startMenu -Filter "*.lnk" -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -in @('Word','Excel','Outlook','PowerPoint','Access','Publisher','OneNote') } |
        Select-Object -ExpandProperty FullName

    if ($shortcuts) {
        foreach ($sc in $shortcuts) { Write-Host "    $sc" }
    } else {
        Write-Host "    Меню Пуск -> Microsoft Office"
    }

    Write-Host ""

    if ($AlreadyInstalled) {
        Write-StatusLine "i" "Повторне встановлення пропущено — Office уже є в системі" "Cyan"
        exit 0
    }

    if ($SetupExitCode -ne 0) {
        if ($officeEntry) {
            Write-StatusLine "!" "setup.exe повернув код $SetupExitCode (в системі: $($officeEntry.Edition) v$($officeEntry.DisplayVersion))" "Yellow"
            exit 3
        }
        Write-StatusLine "X" "setup.exe завершився з помилкою (код $SetupExitCode)" "Red"
        exit 1
    }

    if (-not $officeEntry) { exit 2 }
    exit 0
}

function Invoke-VerifyOfficeRemoved {
    param(
        [int]$SetupExitCode = 0
    )

    Initialize-UkrainianConsole
    $ErrorActionPreference = 'SilentlyContinue'

    Write-Host ""
    Write-StatusLine 'i' 'Перевірка реєстру Windows...' 'Cyan'

    $uninstallKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $officeEntries = Get-ItemProperty $uninstallKeys -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName -match 'Office LTSC|Microsoft 365|Office Professional Plus|Microsoft Office' -and
            $_.DisplayName -notmatch 'Component|Extensibility|Licensing'
        }

    $ctr = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue

    Write-Host ''
    if ($officeEntries) {
        Write-StatusLine '!' 'Office ще присутній у системі:' 'Yellow'
        foreach ($entry in $officeEntries) {
            Write-Host "      - $($entry.DisplayName)"
        }
        if ($SetupExitCode -ne 0) { exit 1 }
        exit 2
    }

    if ($ctr) {
        Write-StatusLine '!' 'Залишкові записи Click-to-Run у реєстрі' 'Yellow'
        exit 2
    }

    Write-StatusLine 'OK' 'Office повністю видалено з системи' 'Green'
    exit 0
}

# --- Teams / Skype check and remove 

function Get-TeamsSkypeStatus {
    $found = @()
    $seen = @{}

    $uninstallKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    foreach ($keyPath in $uninstallKeys) {
        $items = Get-ItemProperty $keyPath -ErrorAction SilentlyContinue | Where-Object {
            $_.DisplayName -and
            $_.DisplayName -match 'Microsoft Teams|Teams Machine-Wide|Skype for Business|^Skype$|SkypeMeetings' -and
            $_.DisplayName -notmatch 'Component|Extensibility|Licensing'
        }

        foreach ($item in $items) {
            $id = "$($item.DisplayName)|Standalone"
            if ($seen.ContainsKey($id)) { continue }
            $seen[$id] = $true

            $found += [PSCustomObject]@{
                Name                 = [string]$item.DisplayName
                Version              = [string]$item.DisplayVersion
                Type                 = 'Standalone'
                UninstallString      = [string]$item.UninstallString
                QuietUninstallString = [string]$item.QuietUninstallString
                Path                 = [string]$item.InstallLocation
            }
        }
    }

    $officeRoots = @(
        'C:\Program Files\Microsoft Office\root\Office16',
        'C:\Program Files (x86)\Microsoft Office\root\Office16',
        'C:\Program Files\Microsoft Office\Office16'
    )

    foreach ($root in $officeRoots) {
        $lync = Join-Path $root 'lync.exe'
        if (Test-Path -LiteralPath $lync) {
            $id = 'Skype for Business (Office)|OfficeComponent'
            if (-not $seen.ContainsKey($id)) {
                $seen[$id] = $true
                $ver = (Get-Item -LiteralPath $lync).VersionInfo.FileVersion
                $found += [PSCustomObject]@{
                    Name                 = 'Skype for Business (Office)'
                    Version              = [string]$ver
                    Type                 = 'OfficeComponent'
                    UninstallString      = ''
                    QuietUninstallString = ''
                    Path                 = $lync
                }
            }
        }
    }

    $teamsUpdate = Join-Path $env:LOCALAPPDATA 'Microsoft\Teams\Update.exe'
    if (Test-Path -LiteralPath $teamsUpdate) {
        $id = 'Microsoft Teams (користувацький)|Standalone'
        if (-not $seen.ContainsKey($id)) {
            $seen[$id] = $true
            $found += [PSCustomObject]@{
                Name                 = 'Microsoft Teams (користувацький)'
                Version              = ''
                Type                 = 'Standalone'
                UninstallString      = "`"$teamsUpdate`" --uninstall -s"
                QuietUninstallString = "`"$teamsUpdate`" --uninstall -s"
                Path                 = $teamsUpdate
            }
        }
    }

    return @($found | Sort-Object Name)
}

function Write-ProgressLine {
    param(
        [string]$Message,
        [TimeSpan]$Elapsed,
        [string]$Suffix = ''
    )

    $elapsedStr = $Elapsed.ToString('mm\:ss')
    $line = "  [>] $($Message)  ($elapsedStr)$Suffix"
    Write-Host ("`r{0,-90}" -f $line) -NoNewline
}

function Wait-ProcessWithProgress {
    param(
        [System.Diagnostics.Process]$Process,
        [string]$Message
    )

    if (-not $Process) { return -1 }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $frames = @('|', '/', '-', '\')
    $frameIndex = 0

    while (-not $Process.HasExited) {
        $frame = $frames[$frameIndex % $frames.Length]
        $frameIndex++
        Write-ProgressLine -Message "$Message $frame" -Elapsed $sw.Elapsed
        Start-Sleep -Milliseconds 400
    }

    Write-Host ""
    return $Process.ExitCode
}

function Invoke-SetupConfigureWithProgress {
    param(
        [string]$SetupPath,
        [string]$ConfigPath,
        [string]$Description
    )

    Write-Host "  [i] $Description" -ForegroundColor Cyan
    Write-Host "      Це може зайняти кілька хвилин..." -ForegroundColor Gray

    $proc = Start-Process -FilePath $SetupPath -ArgumentList @('/configure', $ConfigPath) -PassThru -NoNewWindow
    $exitCode = Wait-ProcessWithProgress -Process $proc -Message 'Office Setup'

    if ($exitCode -eq 0 -or $exitCode -eq 3010) {
        Write-Host "  [OK] setup.exe завершено (код $exitCode)" -ForegroundColor Green
    } else {
        Write-Host "  [!] setup.exe повернув код $exitCode" -ForegroundColor Yellow
    }

    return $exitCode
}

function Wait-RemovalSettle {
    param(
        [scriptblock]$GetRemaining,
        [int]$MaxSeconds = 90
    )

    Write-Host ""
    Write-Host "  [i] Очікування завершення видалення..." -ForegroundColor Cyan

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $lastCount = -1

    while ($sw.Elapsed.TotalSeconds -lt $MaxSeconds) {
        $remaining = @(& $GetRemaining)
        $count = $remaining.Count

        if ($count -eq 0) {
            Write-Host ""
            return @()
        }

        if ($count -ne $lastCount) {
            Write-Host ""
            Write-Host "  [i] Залишилось компонентів: $count" -ForegroundColor Gray
            foreach ($item in $remaining) {
                Write-Host "      - $($item.Name)" -ForegroundColor DarkGray
            }
            $lastCount = $count
        }

        Write-ProgressLine -Message 'Перевірка результату' -Elapsed $sw.Elapsed
        Start-Sleep -Seconds 2
    }

    Write-Host ""
    return @(& $GetRemaining)
}

function Invoke-UninstallRegistryApp {
    param($Item)

    $cmd = if ($Item.QuietUninstallString) { $Item.QuietUninstallString.Trim() }
           elseif ($Item.UninstallString) { $Item.UninstallString.Trim() }
           else { return $false }

    Write-Host "  [i] Запуск деінсталятора: $($Item.Name)" -ForegroundColor Cyan
    Write-Host "      Це може зайняти 1-3 хвилини..." -ForegroundColor Gray

    if ($cmd -match '(?i)msiexec') {
        $args = $cmd -replace '(?i)^(\s*"?)?msiexec\.exe(")?\s*', ''
        $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $args -PassThru -NoNewWindow
        $exitCode = Wait-ProcessWithProgress -Process $proc -Message $Item.Name
        if ($exitCode -eq 0 -or $exitCode -eq 3010) {
            Write-Host "  [OK] $($Item.Name) — видалено" -ForegroundColor Green
            return $true
        }
        Write-Host "  [!] $($Item.Name) — код $exitCode" -ForegroundColor Yellow
        return $false
    }

    if ($cmd -match '^"([^"]+)"\s*(.*)$') {
        $exe = $Matches[1]
        $args = $Matches[2]
    } elseif ($cmd -match '^(\S+\.exe)\s*(.*)$') {
        $exe = $Matches[1]
        $args = $Matches[2]
    } else {
        Write-Host "  [i] Виконання команди видалення..." -ForegroundColor Gray
        $proc = Start-Process -FilePath 'cmd.exe' -ArgumentList @('/c', $cmd) -PassThru -NoNewWindow -Wait
        $ok = ($proc.ExitCode -eq 0)
        if ($ok) {
            Write-Host "  [OK] $($Item.Name) — видалено" -ForegroundColor Green
        } else {
            Write-Host "  [!] $($Item.Name) — код $($proc.ExitCode)" -ForegroundColor Yellow
        }
        return $ok
    }

    if (-not (Test-Path -LiteralPath $exe)) {
        Write-Host "  [!] Файл не знайдено: $exe" -ForegroundColor Yellow
        return $false
    }

    $proc = Start-Process -FilePath $exe -ArgumentList $args -PassThru -NoNewWindow
    $exitCode = Wait-ProcessWithProgress -Process $proc -Message $Item.Name

    if ($exitCode -eq 0) {
        Write-Host "  [OK] $($Item.Name) — видалено" -ForegroundColor Green
        return $true
    }

    Write-Host "  [!] $($Item.Name) — код $exitCode" -ForegroundColor Yellow
    return $false
}

function Invoke-CheckTeamsSkype {
    Initialize-UkrainianConsole
    $ErrorActionPreference = 'SilentlyContinue'

    Write-Host ""
    Write-StatusLine 'i' 'Перевірка Skype та Microsoft Teams...' 'Cyan'
    Write-Host ""

    $items = @(Get-TeamsSkypeStatus)

    if ($items.Count -eq 0) {
        Write-StatusLine 'OK' 'Skype та Teams не знайдено в системі' 'Green'
        exit 0
    }

    Write-StatusLine '!' "Знайдено $($items.Count) компонент(и):" 'Yellow'
    Write-Host ""
    foreach ($item in $items) {
        Write-Host "    - $($item.Name)"
        if ($item.Version) { Write-Host "      Версія: $($item.Version)" }
        Write-Host "      Тип:    $($item.Type)"
        if ($item.Path) { Write-Host "      Шлях:   $($item.Path)" }
        Write-Host ""
    }

    exit 2
}

function Invoke-RemoveTeamsSkype {
    param([string]$InstallDir)

    Initialize-UkrainianConsole
    $ErrorActionPreference = 'Continue'

    if (-not $InstallDir) {
        $paths = Resolve-OfficePaths -BasePath $PSScriptRoot
        $InstallDir = $paths.RootDir
    }

    Write-Host ""
    Write-StatusLine 'i' 'Видалення Skype та Microsoft Teams...' 'Cyan'
    Write-Host ""

    $before = @(Get-TeamsSkypeStatus)
    if ($before.Count -eq 0) {
        Write-StatusLine 'OK' 'Skype та Teams не знайдено — нічого видаляти' 'Green'
        exit 0
    }

    $needsOfficeConfigure = @($before | Where-Object { $_.Type -eq 'OfficeComponent' })
    $standalone = @($before | Where-Object { $_.Type -eq 'Standalone' })
    $setupAvailable = Test-Path -LiteralPath (Join-Path $InstallDir 'setup.exe')
    $totalSteps = 0
    if ($needsOfficeConfigure -and $setupAvailable) { $totalSteps++ }
    if ($standalone.Count -gt 0) { $totalSteps += $standalone.Count }
    if ($totalSteps -eq 0 -and $needsOfficeConfigure) { $totalSteps = 1 }

    Write-StatusLine '!' "Знайдено $($before.Count) компонент(и) — план видалення:" 'Yellow'
    Write-Host ""
    $planStep = 0
    if ($needsOfficeConfigure) {
        $planStep++
        $via = if ($setupAvailable) { 'setup.exe (Office)' } else { 'пропуск — немає setup.exe' }
        Write-Host "    [$planStep] Skype for Business — $via"
    }
    foreach ($item in $standalone) {
        $planStep++
        Write-Host "    [$planStep] $($item.Name)"
    }
    Write-Host ""

    $currentStep = 0

    if ($needsOfficeConfigure -and $setupAvailable) {
        $currentStep++
        Write-Host "  ── Крок $currentStep/$($totalSteps): Skype for Business (Office) ──" -ForegroundColor White
        Invoke-PrepareOfficeConfig -InstallDir $InstallDir -Mode RemoveApps | Out-Null
        $config = Join-Path $InstallDir 'configuration-remove-teams-skype.xml'
        Set-Location $InstallDir
        $null = Invoke-SetupConfigureWithProgress -SetupPath (Join-Path $InstallDir 'setup.exe') `
            -ConfigPath $config -Description 'Видалення Skype for Business з пакету Office'
        Write-Host ""
    } elseif ($needsOfficeConfigure) {
        Write-Host "  [!] setup.exe не знайдено — пропуск видалення Skype for Business (пункт [1])" -ForegroundColor Yellow
        Write-Host ""
    }

    foreach ($item in $standalone) {
        $currentStep++
        Write-Host "  ── Крок $currentStep/$($totalSteps): $($item.Name) ──" -ForegroundColor White
        $null = Invoke-UninstallRegistryApp -Item $item
        Write-Host ""
    }

    $after = Wait-RemovalSettle -GetRemaining { Get-TeamsSkypeStatus }
    Write-Host ""

    if ($after.Count -eq 0) {
        Write-StatusLine 'OK' 'Skype та Teams успішно видалено' 'Green'
        exit 0
    }

    Write-StatusLine '!' 'Залишились компоненти:' 'Yellow'
    foreach ($item in $after) {
        Write-Host "      - $($item.Name)"
    }
    Write-Host ""
    Write-Host "  Спробуйте перезавантажити ПК або пункт [4] — перевстановити Office." -ForegroundColor Yellow
    exit 1
}

# --- Dispatcher 

switch ($Command) {
    'ResolvePaths' {
        Resolve-OfficePaths -BasePath $(if ($BasePath) { $BasePath } else { $PSScriptRoot })
    }
    'PrepareConfig' {
        Invoke-PrepareOfficeConfig -InstallDir $InstallDir -Mode $Mode
    }
    'PrepareInstallConfig' {
        Invoke-PrepareInstallConfig -InstallDir $InstallDir -Reinstall:$Reinstall
    }
    'Download' {
        Invoke-DownloadOfficePackage -InstallDir $InstallDir -ScriptDir $ScriptDir
    }
    'DownloadOdt' {
        Invoke-DownloadOdt -TargetDir $InstallDir
    }
    'Install' {
        Invoke-InstallOffice2021
    }
    'VerifyInstall' {
        Invoke-VerifyOfficeInstall -SetupExitCode $SetupExitCode -AlreadyInstalled:$AlreadyInstalled
    }
    'VerifyRemoved' {
        Invoke-VerifyOfficeRemoved -SetupExitCode $SetupExitCode
    }
    'GetInstalledOffice' {
        Invoke-GetInstalledOffice -Format $OutputFormat
    }
    'CheckTeamsSkype' {
        Invoke-CheckTeamsSkype
    }
    'RemoveTeamsSkype' {
        Invoke-RemoveTeamsSkype -InstallDir $InstallDir
    }
}
