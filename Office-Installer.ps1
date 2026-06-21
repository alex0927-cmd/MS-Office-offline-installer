# Office 2021 Professional Plus — об'єднаний інсталятор
# Замінює: Init-ConsoleEncoding, Resolve-OfficePaths, Prepare-OfficeConfig,
#         Prepare-InstallConfig, Download-OfficePackage, Install-Office2021,
#         Verify-OfficeInstall, Verify-OfficeRemoved

param(
    [ValidateSet('PrepareConfig', 'PrepareInstallConfig', 'Download', 'DownloadOdt', 'Install', 'Uninstall', 'VerifyInstall', 'VerifyRemoved', 'ResolvePaths', 'GetInstalledOffice', 'CheckTeamsSkype', 'RemoveTeamsSkype', 'Activate')]
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
    [switch]$AlreadyInstalled,

    [string]$ProductKey,
    [string]$KmsHost,
    [int]$KmsPort = 0,
    [switch]$ActivateOnly
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

function Get-LocalOfficePackageVersion {
    param([string]$InstallDir)

    $dataPath = Join-Path $InstallDir 'Office\Data'
    if (-not (Test-Path -LiteralPath $dataPath)) { return $null }

    $versionDir = Get-ChildItem -LiteralPath $dataPath -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' } |
        Sort-Object Name -Descending |
        Select-Object -First 1

    if ($versionDir) { return $versionDir.Name }

    $cab = Get-ChildItem -LiteralPath $dataPath -Filter 'v64_*.cab' -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($cab -and $cab.Name -match 'v64_(.+)\.cab') { return $Matches[1] }

    return $null
}

function Get-InstallDisplayLevel {
    param([string]$InstallDir)

    foreach ($searchDir in @($InstallDir, $PSScriptRoot)) {
        if (-not $searchDir) { continue }
        $configFile = Join-Path $searchDir 'configurations.xml'
        if (-not (Test-Path -LiteralPath $configFile)) { continue }

        [xml]$doc = Get-Content -LiteralPath $configFile
        if ($doc.OfficeConfigurations.InstallProfile.DisplayLevel) {
            return [string]$doc.OfficeConfigurations.InstallProfile.DisplayLevel
        }
    }

    return 'None'
}

function Add-OdtOfflineCommonOptions {
    param(
        [System.Xml.XmlDocument]$Doc,
        [System.Xml.XmlElement]$Cfg
    )

    $updates = $Doc.CreateElement('Updates')
    [void]$updates.SetAttribute('Enabled', 'FALSE')
    [void]$Cfg.AppendChild($updates)

    foreach ($prop in @(
            @{ Name = 'FORCEAPPSHUTDOWN'; Value = 'TRUE' },
            @{ Name = 'SharedComputerLicensing'; Value = '0' }
        )) {
        $property = $Doc.CreateElement('Property')
        [void]$property.SetAttribute('Name', $prop.Name)
        [void]$property.SetAttribute('Value', $prop.Value)
        [void]$Cfg.AppendChild($property)
    }
}

function Add-OdtOfflineInstallOptions {
    param(
        [System.Xml.XmlDocument]$Doc,
        [System.Xml.XmlElement]$Cfg,
        [System.Xml.XmlElement]$Add,
        [string]$SourcePath,
        [string]$Version
    )

    [void]$Add.SetAttribute('SourcePath', $SourcePath)
    [void]$Add.SetAttribute('AllowCdnFallback', 'False')
    if ($Version) {
        [void]$Add.SetAttribute('Version', $Version)
    }

    Add-OdtOfflineCommonOptions -Doc $Doc -Cfg $Cfg
}

function Add-OdtOfflineRemoveOptions {
    param(
        [System.Xml.XmlDocument]$Doc,
        [System.Xml.XmlElement]$Cfg,
        [System.Xml.XmlElement]$Remove,
        [string]$InstallDir
    )

    $sourcePath = (Resolve-Path -LiteralPath $InstallDir).Path
    [void]$Remove.SetAttribute('SourcePath', $sourcePath)

    Add-OdtOfflineCommonOptions -Doc $Doc -Cfg $Cfg
}

function Add-OdtInstallLoggingNode {
    param(
        [System.Xml.XmlDocument]$Doc,
        [System.Xml.XmlElement]$Cfg,
        [string]$LogDir
    )

    if (-not (Test-Path -LiteralPath $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    }

    $logging = $Doc.CreateElement('Logging')
    [void]$logging.SetAttribute('Level', 'Standard')
    [void]$logging.SetAttribute('Path', $LogDir)
    [void]$Cfg.AppendChild($logging)
}

function Test-OfflineOfficePackage {
    param(
        [string]$InstallDir,
        [switch]$Quiet
    )

    $issues = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]
    $dataPath = Join-Path $InstallDir 'Office\Data'

    if (-not (Test-Path -LiteralPath (Join-Path $InstallDir 'setup.exe'))) {
        $issues.Add('setup.exe не знайдено — виконайте пункт [1]')
    }

    if (-not (Test-Path -LiteralPath $dataPath)) {
        $issues.Add('Office\Data не знайдено — виконайте пункт [2]')
    } else {
        $officeSize = Get-PackageSize -Path (Join-Path $InstallDir 'Office')
        if ($officeSize -lt 1GB) {
            $issues.Add("Офлайн-пакет занадто малий ($(Format-Size $officeSize), потрібно ~2 GB+) — повторіть пункт [2] на ПК з інтернетом")
        } elseif ($officeSize -lt 1.5GB) {
            $warnings.Add("Розмір пакету підозріло малий ($(Format-Size $officeSize)) — повне завантаження зазвичай ~2 GB+")
        } elseif (-not $Quiet) {
            Write-Host "  [OK] Розмір офлайн-пакету: $(Format-Size $officeSize)" -ForegroundColor Green
        }

        $version = Get-LocalOfficePackageVersion -InstallDir $InstallDir
        if (-not $version) {
            $issues.Add('Не вдалося визначити версію пакету в Office\Data')
        } elseif (-not $Quiet) {
            Write-Host "  [OK] Версія офлайн-пакету: $version" -ForegroundColor Green
        }

        $requiredPatterns = @(
            @{ Label = 'v64*.cab'; Pattern = 'v64*.cab' }
            @{ Label = 'i640.cab'; Pattern = 'i640.cab' }
            @{ Label = 'stream.x64.*.dat'; Pattern = 'stream.x64.*.dat' }
        )

        foreach ($item in $requiredPatterns) {
            $found = @(Get-ChildItem -LiteralPath $dataPath -Recurse -Filter $item.Pattern -File -ErrorAction SilentlyContinue)
            if ($found.Count -eq 0) {
                $issues.Add("Відсутній обов'язковий файл: $($item.Label)")
            } elseif (-not $Quiet) {
                $largest = ($found | Sort-Object Length -Descending | Select-Object -First 1)
                Write-Host "  [OK] $($item.Label) — $(Format-Size $largest.Length)" -ForegroundColor Green
            }
        }

        $streamFiles = @(Get-ChildItem -LiteralPath $dataPath -Recurse -Filter 'stream.x64.*.dat' -File -ErrorAction SilentlyContinue)
        if ($streamFiles.Count -gt 0) {
            $streamTotal = ($streamFiles | Measure-Object Length -Sum).Sum
            if ($streamTotal -lt 500MB) {
                $issues.Add("Файли stream.x64.*.dat занадто малі ($(Format-Size $streamTotal)) — офлайн-пакет завантажено не повністю")
            }
        }
    }

    if ($warnings.Count -gt 0 -and -not $Quiet) {
        foreach ($warning in $warnings) {
            Write-Host "  [!] $warning" -ForegroundColor Yellow
        }
    }

    if ($issues.Count -gt 0) {
        if (-not $Quiet) {
            Write-Host "  [X] Офлайн-пакет не готовий до встановлення:" -ForegroundColor Red
            foreach ($issue in $issues) {
                Write-Host "      - $issue" -ForegroundColor Yellow
            }
        }
        return $false
    }

    return $true
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
            $displayLevel = Get-InstallDisplayLevel -InstallDir $InstallDir
            $logDir = Join-Path $InstallDir 'Logs'

            $remove = $doc.CreateElement('Remove')
            [void]$remove.SetAttribute('All', 'TRUE')

            Add-OdtOfflineRemoveOptions -Doc $doc -Cfg $cfg -Remove $remove -InstallDir $InstallDir
            Add-OdtInstallLoggingNode -Doc $doc -Cfg $cfg -LogDir $logDir
            [void]$cfg.AppendChild($remove)

            $display = $doc.CreateElement('Display')
            [void]$display.SetAttribute('Level', $displayLevel)
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

            $sourcePath = (Resolve-Path -LiteralPath $InstallDir).Path
            $packageVersion = Get-LocalOfficePackageVersion -InstallDir $InstallDir
            $displayLevel = Get-InstallDisplayLevel -InstallDir $InstallDir
            $logDir = Join-Path $InstallDir 'Logs'

            $add = $doc.CreateElement('Add')
            [void]$add.SetAttribute('OfficeClientEdition', $src.Add.OfficeClientEdition)
            [void]$add.SetAttribute('Channel', $src.Add.Channel)

            Add-OdtOfflineInstallOptions -Doc $doc -Cfg $cfg -Add $add `
                -SourcePath $sourcePath -Version $packageVersion
            Add-OdtInstallLoggingNode -Doc $doc -Cfg $cfg -LogDir $logDir

            Add-OdtProductNode -Doc $doc -AddElement $add -SrcProduct $src.Add.Product
            [void]$cfg.AppendChild($add)

            $display = $doc.CreateElement('Display')
            [void]$display.SetAttribute('Level', $displayLevel)
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
    $sourcePath = (Resolve-Path -LiteralPath $InstallDir).Path
    $packageVersion = Get-LocalOfficePackageVersion -InstallDir $InstallDir
    $displayLevel = Get-InstallDisplayLevel -InstallDir $InstallDir

    if ($cfg.Add) {
        $cfg.Add.SourcePath = $sourcePath
        $cfg.Add.AllowCdnFallback = 'False'
        if ($packageVersion) {
            $cfg.Add.Version = $packageVersion
        }
    }

    if (-not $cfg.Updates) {
        $updates = $doc.CreateElement('Updates')
        $updates.SetAttribute('Enabled', 'FALSE')
        [void]$cfg.AppendChild($updates)
    } else {
        $cfg.Updates.Enabled = 'FALSE'
    }

    $cfg.Display.Level = $displayLevel
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

# --- Download helpers

function Enable-ModernTls {
    try {
        $protocols = [Net.ServicePointManager]::SecurityProtocol
        $protocols = $protocols -bor [Net.SecurityProtocolType]::Tls12
        if ([Enum]::IsDefined([Net.SecurityProtocolType], 'Tls13')) {
            $protocols = $protocols -bor [Net.SecurityProtocolType]::Tls13
        }
        [Net.ServicePointManager]::SecurityProtocol = $protocols
    } catch { }
}

function Test-MicrosoftDownloadAccess {
    param(
        [string]$Url = 'https://download.microsoft.com'
    )

    Enable-ModernTls

    $hostName = 'download.microsoft.com'
    try {
        $hostName = ([Uri]$Url).Host
    } catch { }

    try {
        $tcp = Test-NetConnection -ComputerName $hostName -Port 443 `
            -WarningAction SilentlyContinue -ErrorAction Stop
        if ($tcp.TcpTestSucceeded) { return $true }
    } catch { }

    try {
        return Test-Connection -ComputerName $hostName -Count 1 -Quiet -ErrorAction Stop
    } catch {
        return $false
    }
}

function Save-RemoteFile {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Destination,
        [int]$TimeoutSec = 300,
        [long]$MinBytes = 100KB
    )

    Enable-ModernTls
    $errors = New-Object System.Collections.Generic.List[string]

    if (Test-Path -LiteralPath $Destination) {
        Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
    }

    try {
        $iwrParams = @{
            Uri             = $Uri
            OutFile         = $Destination
            UseBasicParsing = $true
            TimeoutSec      = $TimeoutSec
        }
        if ($PSVersionTable.PSVersion.Major -lt 6) {
            $iwrParams['Headers'] = @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
        }
        Invoke-WebRequest @iwrParams -ErrorAction Stop
        if ((Test-Path -LiteralPath $Destination) -and (Get-Item -LiteralPath $Destination).Length -ge $MinBytes) {
            return
        }
        Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
        throw 'Завантажений файл занадто малий або порожній'
    } catch {
        $errors.Add("Invoke-WebRequest: $($_.Exception.Message)")
    }

    try {
        if (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue) {
            Start-BitsTransfer -Source $Uri -Destination $Destination -TransferType Download -ErrorAction Stop
            if ((Test-Path -LiteralPath $Destination) -and (Get-Item -LiteralPath $Destination).Length -ge $MinBytes) {
                return
            }
            Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
            throw 'Завантажений файл занадто малий або порожній'
        }
    } catch {
        $errors.Add("BITS: $($_.Exception.Message)")
    }

    try {
        if (Get-Command curl.exe -ErrorAction SilentlyContinue) {
            & curl.exe -fL --retry 3 --connect-timeout 30 --max-time $TimeoutSec `
                -o $Destination $Uri 2>&1 | Out-Null
            if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $Destination) `
                -and (Get-Item -LiteralPath $Destination).Length -ge $MinBytes) {
                return
            }
            Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
            throw "curl завершився з кодом $LASTEXITCODE"
        }
    } catch {
        $errors.Add("curl: $($_.Exception.Message)")
    }

    $details = ($errors | ForEach-Object { "      - $_" }) -join [Environment]::NewLine
    throw "Не вдалося завантажити файл з Microsoft.`n$details"
}

function Invoke-OdtExtract {
    param(
        [Parameter(Mandatory)][string]$InstallerPath,
        [Parameter(Mandatory)][string]$TargetDir
    )

    $extractArgList = @("/extract:`"$TargetDir`"", '/quiet')
    $cmdLine = "`"$InstallerPath`" /extract:`"$TargetDir`" /quiet"

    cmd /c $cmdLine 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) { return 0 }

    try {
        $proc = Start-Process -FilePath $InstallerPath -ArgumentList $extractArgList `
            -Wait -PassThru -NoNewWindow -ErrorAction Stop
        return $proc.ExitCode
    } catch {
        if ($_.Exception.Message -notmatch 'elevation|administrator|адміністратор') {
            return 1
        }

        try {
            $proc = Start-Process -FilePath $InstallerPath -ArgumentList $extractArgList `
                -Wait -PassThru -Verb RunAs -ErrorAction Stop
            return $proc.ExitCode
        } catch {
            return 1
        }
    }
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
    $url = Get-OdtDownloadUrl -Dir $TargetDir

    Write-Host ""
    Write-Host "  [i] Перевірка доступу до download.microsoft.com (HTTPS)..." -ForegroundColor Cyan

    if (-not (Test-MicrosoftDownloadAccess -Url $url)) {
        Write-Host "  [X] Немає доступу до серверів Microsoft (порт 443)." -ForegroundColor Red
        Write-Host "      Перевірте інтернет, проксі, VPN або брандмауер." -ForegroundColor Yellow
        exit 1
    }
    Write-Host "  [OK] Сервер Microsoft доступний" -ForegroundColor Green

    if (Test-Path -LiteralPath $setupPath) {
        Write-Host ""
        Write-Host "  [i] setup.exe уже існує у $TargetDir" -ForegroundColor Yellow
    }

    $tempExe = Join-Path $env:TEMP ("odt_{0}.exe" -f ([guid]::NewGuid().ToString('N').Substring(0, 8)))

    Write-Host ""
    Write-Host "  Завантаження Office Deployment Tool..." -ForegroundColor Cyan
    Write-Host "  Джерело: Microsoft" -ForegroundColor Gray
    Write-Host "  URL:     $url" -ForegroundColor Gray
    Write-Host "  Куди:    $TargetDir" -ForegroundColor Gray
    Write-Host ""

    try {
        Save-RemoteFile -Uri $url -Destination $tempExe -MinBytes 1MB
    } catch {
        Write-Host "  [X] Не вдалося завантажити ODT:" -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Write-Host "  Поради:" -ForegroundColor Yellow
        Write-Host "    - завантажте ODT вручну з https://www.microsoft.com/download/details.aspx?id=49117" -ForegroundColor Yellow
        Write-Host "    - запустіть officedeploymenttool_*.exe і вкажіть папку: $TargetDir" -ForegroundColor Yellow
        exit 1
    }

    Write-Host "  [OK] ODT завантажено ($(Format-Size (Get-Item -LiteralPath $tempExe).Length))" -ForegroundColor Green
    Write-Host "  Розпакування setup.exe..." -ForegroundColor Cyan

    $extractCode = Invoke-OdtExtract -InstallerPath $tempExe -TargetDir $TargetDir

    Remove-Item -LiteralPath $tempExe -Force -ErrorAction SilentlyContinue

    Write-Host ""

    if (Test-Path -LiteralPath $setupPath) {
        if ($extractCode -and $extractCode -ne 0) {
            Write-Host "  [i] setup.exe створено (код розпакування: $extractCode)" -ForegroundColor Yellow
        }
        Write-Host "  [OK] setup.exe готовий до роботи" -ForegroundColor Green
        Write-Host "       $setupPath" -ForegroundColor Green
        exit 0
    }

    Write-Host "  [X] setup.exe не з'явився (код розпакування: $extractCode)" -ForegroundColor Red
    Write-Host "      Спробуйте запустити інсталятор ODT вручну і вказати папку $TargetDir" -ForegroundColor Yellow
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
    Write-Host "  [i] Перевірка доступу до серверів Microsoft (HTTPS)..." -ForegroundColor Cyan

    if (-not (Test-MicrosoftDownloadAccess)) {
        Write-Host "  [X] Немає доступу до серверів Microsoft (порт 443)." -ForegroundColor Red
        Write-Host "      Перевірте інтернет, проксі, VPN або брандмауер." -ForegroundColor Yellow
        exit 1
    }
    Write-Host "  [OK] Сервер Microsoft доступний" -ForegroundColor Green

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
    $logDir = Get-OfficeInstallerLogDirectory -InstallDir $InstallDir
    $workDir = Get-OfficeSetupWorkDirectory -LogDirectory $logDir

    $job = Start-Job -ScriptBlock {
        param($SetupPath, $ConfigPath, $WorkDir, $OdtWorkDir)
        Set-Location $WorkDir
        if ($OdtWorkDir) {
            $env:TEMP = $OdtWorkDir
            $env:TMP = $OdtWorkDir
        }
        & $SetupPath /download $ConfigPath
        return $LASTEXITCODE
    } -ArgumentList $setup, $config, $InstallDir, $workDir

    $lastSize = 0L
    $lastTick = Get-Date
    $stableSize = 0L
    $stableSince = Get-Date
    $lastLogSync = [DateTime]::MinValue

    while ($job.State -eq 'Running') {
        Start-Sleep -Seconds 2

        if ($logDir -and ((Get-Date) - $lastLogSync).TotalSeconds -ge 5) {
            Sync-OfficeSetupLogsToDirectory -LogDirectory $logDir -NotBefore $startTime | Out-Null
            $lastLogSync = Get-Date
        }

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

    if ($logDir) {
        Sync-OfficeSetupLogsToDirectory -LogDirectory $logDir -NotBefore $startTime | Out-Null
        Write-OfficeSetupLogSummary -LogDirectory $logDir -NotBefore $startTime
    }

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
    Write-Host "      Перевірте логи у папці Logs" -ForegroundColor Yellow
    exit 1
}

# --- Install-Office2021 

function Invoke-InstallOffice2021 {
    Initialize-UkrainianConsole
    $ErrorActionPreference = 'Stop'

    if ($InstallDir) {
        $root = (Resolve-Path -LiteralPath $InstallDir).Path
    } else {
        $paths = Resolve-OfficePaths -BasePath $PSScriptRoot
        $root = $paths.RootDir
    }

    if (-not (Test-Path -LiteralPath (Join-Path $root 'setup.exe'))) {
        Write-Host "setup.exe не знайдено. Спочатку виконайте пункт [1] у меню." -ForegroundColor Red
        exit 1
    }

    Write-Host ""
    Write-Host "  [i] Перевірка офлайн-пакету..." -ForegroundColor Cyan
    if (-not (Test-OfflineOfficePackage -InstallDir $root)) {
        exit 1
    }

    Set-Location $root

    $setupMode = if ($Reinstall -or $Mode -eq 'Reinstall') { 'Reinstall' } else { 'Install' }
    if ($setupMode -eq 'Reinstall') {
        Write-Host ""
        Write-Host "  [i] Підготовка конфігурації (перевстановлення)..." -ForegroundColor Cyan
    } else {
        Write-Host ""
        Write-Host "  [i] Підготовка конфігурації встановлення..." -ForegroundColor Cyan
    }

    Prepare-OfficeInstallEnvironment -InstallDir $root -Reinstall:($setupMode -eq 'Reinstall')

    Invoke-PrepareOfficeConfig -InstallDir $root -Mode $setupMode | Out-Null

    $packageVersion = Get-LocalOfficePackageVersion -InstallDir $root
    $displayLevel = Get-InstallDisplayLevel -InstallDir $root
    $configPath = Join-Path $root 'configuration-install.xml'
    $setupPath = Join-Path $root 'setup.exe'
    $showTerminalProgress = ($displayLevel -eq 'None')

    Write-Host ""
    Write-Host "  Встановлення Office 2021 Pro Plus (офлайн)..." -ForegroundColor Cyan
    Write-Host "  Пакет:   $root" -ForegroundColor Gray
    Write-Host "  Версія:  $packageVersion" -ForegroundColor Gray
    Write-Host "  Режим:   без звернення до інтернету (CDN вимкнено)" -ForegroundColor Gray
    if ($showTerminalProgress) {
        Write-Host "  UI:      прихований — прогрес у цьому вікні" -ForegroundColor Gray
    }
    Write-Host ""

    $exitCode = Invoke-SetupConfigureWithProgress `
        -SetupPath $setupPath `
        -ConfigPath $configPath `
        -InstallDir $root `
        -ShowTerminalProgress:$showTerminalProgress `
        -SetupOperation Install `
        -Description $(if ($setupMode -eq 'Reinstall') { 'Перевстановлення Office з офлайн-пакету' } else { 'Встановлення Office з офлайн-пакету' })

    Write-Host ""
    if ($exitCode -eq 0 -or $exitCode -eq 3010) {
        Write-Host "Office 2021 Pro Plus встановлено успішно!" -ForegroundColor Green
        exit 0
    }

    Write-Host "Помилка встановлення. Код: $exitCode" -ForegroundColor Red
    $help = Get-OfficeSetupErrorHelp -ExitCode $exitCode
    if ($help) {
        Write-Host $help -ForegroundColor Yellow
    } else {
        Write-Host "Перевірте лог у папці Logs" -ForegroundColor Yellow
    }
    exit 1
}

function Invoke-UninstallOffice {
    Initialize-UkrainianConsole
    $ErrorActionPreference = 'Stop'

    if ($InstallDir) {
        $root = (Resolve-Path -LiteralPath $InstallDir).Path
    } else {
        $paths = Resolve-OfficePaths -BasePath $PSScriptRoot
        $root = $paths.RootDir
    }

    if (-not (Test-Path -LiteralPath (Join-Path $root 'setup.exe'))) {
        Write-Host "setup.exe не знайдено. Спочатку виконайте пункт [1] у меню." -ForegroundColor Red
        exit 1
    }

    $officeItems = @(Get-InstalledOfficeInfo)
    if ($officeItems.Count -eq 0) {
        Write-Host ""
        Write-Host "  [i] Office у системі не знайдено — видаляти нічого." -ForegroundColor Yellow
        exit 0
    }

    Set-Location $root

    Write-Host ""
    Write-Host "  [i] Підготовка до видалення..." -ForegroundColor Cyan
    Stop-OfficeSetupProcesses
    Start-OfficeClickToRunService | Out-Null

    Invoke-PrepareOfficeConfig -InstallDir $root -Mode Uninstall | Out-Null

    $displayLevel = Get-InstallDisplayLevel -InstallDir $root
    $configPath = Join-Path $root 'configuration-remove.xml'
    $setupPath = Join-Path $root 'setup.exe'
    $showTerminalProgress = ($displayLevel -eq 'None')

    Write-Host ""
    Write-Host "  Видалення Microsoft Office..." -ForegroundColor Cyan
    foreach ($item in $officeItems) {
        Write-Host "  Знайдено: $($item.DisplayName) v$($item.DisplayVersion)" -ForegroundColor Gray
    }
    Write-Host "  Режим:   без вікна Microsoft, без оновлень з інтернету" -ForegroundColor Gray
    if ($showTerminalProgress) {
        Write-Host "  UI:      прихований — прогрес у цьому вікні" -ForegroundColor Gray
    }
    Write-Host ""

    $exitCode = Invoke-SetupConfigureWithProgress `
        -SetupPath $setupPath `
        -ConfigPath $configPath `
        -InstallDir $root `
        -ShowTerminalProgress:$showTerminalProgress `
        -SetupOperation Uninstall `
        -Description 'Видалення Office'

    Write-Host ""
    if ($exitCode -eq 0 -or $exitCode -eq 3010) {
        Write-Host "Office видалено успішно!" -ForegroundColor Green
        exit 0
    }

    Write-Host "Помилка видалення. Код: $exitCode" -ForegroundColor Red
    $help = Get-OfficeSetupErrorHelp -ExitCode $exitCode
    if ($help) {
        Write-Host $help -ForegroundColor Yellow
    } else {
        Write-Host "Перевірте лог у папці Logs" -ForegroundColor Yellow
    }
    exit 1
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

function Get-OsppScriptPath {
    $candidates = @(
        'C:\Program Files\Microsoft Office\Office16\OSPP.VBS',
        'C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS',
        'C:\Program Files\Microsoft Office\root\Office16\OSPP.VBS'
    )
    foreach ($root in @(
        'C:\Program Files\Microsoft Office\root\Office16',
        'C:\Program Files\Microsoft Office\Office16'
    )) {
        $candidates += Join-Path $root 'OSPP.VBS'
    }
    foreach ($path in ($candidates | Select-Object -Unique)) {
        if (Test-Path -LiteralPath $path) { return (Resolve-Path -LiteralPath $path).Path }
    }
    return $null
}

function Test-WindowsScriptHostEnabled {
    foreach ($regPath in @(
        'HKLM:\Software\Microsoft\Windows Script Host\Settings',
        'HKCU:\Software\Microsoft\Windows Script Host\Settings'
    )) {
        if (-not (Test-Path -LiteralPath $regPath)) { continue }
        $enabled = (Get-ItemProperty -LiteralPath $regPath -Name Enabled -ErrorAction SilentlyContinue).Enabled
        if ($null -ne $enabled -and [int]$enabled -eq 0) { return $false }
    }
    return $true
}

function Test-OsppVbsBlockedOutput {
    param([string]$Text)
    return ($Text -match '(?i)(Script Host access is disabled|execution of scripts is disabled|ActiveX component can''t create object)')
}

function Test-OfficeLicenseUsesWmi {
    return -not (Test-WindowsScriptHostEnabled)
}

function Get-OfficeSoftwareLicensingProducts {
    $appId = '59A52893-A989-479D-AF46-F275C0160663'
    return @(Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction SilentlyContinue |
        Where-Object { $_.ApplicationId -eq $appId -and $_.PartialProductKey })
}

function Convert-OfficeLicenseStatusCode {
    param([int]$Code)
    switch ($Code) {
        0 { return '---UNLICENSED---' }
        1 { return '---LICENSED---' }
        5 { return '---NOTIFICATIONS---' }
        2 { return '---OOB_GRACE---' }
        3 { return '---OOT_GRACE---' }
        default { return "---STATUS_${Code}---" }
    }
}

function Get-OfficeLicenseInfoFromWmi {
    $products = @(Get-OfficeSoftwareLicensingProducts)
    if ($products.Count -eq 0) {
        return [PSCustomObject]@{
            Available = $true; TimedOut = $false; Licensed = $false
            LicenseStatus = $null; LicenseDescription = $null; LicenseName = $null
            ProductKeyLast5 = $null; KmsHost = $null
            Error = 'Ключ продукту не встановлено'
            Raw = $null; Backend = 'Wmi'
        }
    }

    $product = $products | Select-Object -First 1
    $status = Convert-OfficeLicenseStatusCode -Code [int]$product.LicenseStatus
    $kmsHost = $null
    if ($product.KeyManagementServiceMachine) { $kmsHost = [string]$product.KeyManagementServiceMachine }

    return [PSCustomObject]@{
        Available = $true; TimedOut = $false; Licensed = ($product.LicenseStatus -eq 1)
        LicenseStatus = $status
        LicenseDescription = [string]$product.Description
        LicenseName = [string]$product.Name
        ProductKeyLast5 = [string]$product.PartialProductKey
        KmsHost = $kmsHost
        Error = $null; Raw = $null; Backend = 'Wmi'
    }
}

function Invoke-OfficeOnlineActivationWmi {
    Write-Host ""
    Write-StatusLine 'i' 'Активація Office (PowerShell / WMI)...' 'Cyan'
    Write-Host "      Потрібен доступ до серверів Microsoft або KMS." -ForegroundColor Gray

    $products = @(Get-OfficeSoftwareLicensingProducts)
    if ($products.Count -eq 0) {
        Write-StatusLine 'X' 'Ключ продукту не встановлено — спочатку введіть ключ' 'Red'
        return 1
    }

    foreach ($product in $products) {
        try {
            $null = Invoke-CimMethod -InputObject $product -MethodName Activate -ErrorAction Stop
        } catch {
            Write-StatusLine '!' "Activate: $($_.Exception.Message)" 'Yellow'
        }
    }

    Start-Sleep -Seconds 2
    $info = Get-OfficeLicenseInfoFromWmi
    if ($info.Licensed) {
        Write-StatusLine 'OK' 'Office активовано успішно' 'Green'
        return 0
    }
    Write-StatusLine '!' 'Office ще не активовано' 'Yellow'
    return 1
}

function Invoke-OfficeProductKeyActivationWmi {
    param([Parameter(Mandatory)][string]$ProductKey)

    $key = Normalize-ProductKey -Key $ProductKey
    if (-not (Test-ProductKeyFormat -Key $key)) {
        Write-StatusLine 'X' 'Невірний формат ключа' 'Red'
        return 2
    }

    Write-Host ""
    Write-StatusLine 'i' 'Встановлення ключа продукту (PowerShell / WMI)...' 'Cyan'
    try {
        $svc = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction Stop
        $result = Invoke-CimMethod -InputObject $svc -MethodName InstallProductKey -Arguments @{ ProductKey = $key } -ErrorAction Stop
        if ($result.ReturnValue -and [int]$result.ReturnValue -ne 0) {
            Write-StatusLine 'X' "InstallProductKey: код $($result.ReturnValue)" 'Red'
            return 1
        }
        $null = Invoke-CimMethod -InputObject $svc -MethodName RefreshLicenseStatus -ErrorAction SilentlyContinue
    } catch {
        Write-StatusLine 'X' $_.Exception.Message 'Red'
        return 1
    }

    Write-StatusLine 'OK' 'Ключ встановлено' 'Green'
    return (Invoke-OfficeOnlineActivationWmi)
}

function Invoke-OfficeKmsActivationWmi {
    param(
        [Parameter(Mandatory)][string]$KmsHost,
        [int]$KmsPort = 1688
    )

    $hostName = $KmsHost.Trim()
    if ([string]::IsNullOrWhiteSpace($hostName)) {
        Write-StatusLine 'X' 'KMS-сервер не вказано' 'Red'
        return 2
    }
    if ($KmsPort -le 0) { $KmsPort = 1688 }

    Write-Host ""
    Write-StatusLine 'i' "Налаштування KMS: ${hostName}:${KmsPort} (PowerShell / WMI)" 'Cyan'
    try {
        $svc = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction Stop
        $null = Invoke-CimMethod -InputObject $svc -MethodName SetKeyManagementServiceMachine -Arguments @{
            KeyManagementServiceMachine = $hostName
        } -ErrorAction Stop
        if ($KmsPort -ne 1688) {
            $null = Invoke-CimMethod -InputObject $svc -MethodName SetKeyManagementServicePort -Arguments @{
                KeyManagementServicePort = $KmsPort
            } -ErrorAction Stop
        }
        $null = Invoke-CimMethod -InputObject $svc -MethodName RefreshLicenseStatus -ErrorAction SilentlyContinue
    } catch {
        Write-StatusLine 'X' $_.Exception.Message 'Red'
        return 1
    }

    return (Invoke-OfficeOnlineActivationWmi)
}

function Invoke-OsppScript {
    param(
        [Parameter(Mandatory)][string]$OsppPath,
        [Parameter(Mandatory)][string[]]$Arguments,
        [int]$TimeoutSeconds = 90
    )

    if (-not (Test-Path -LiteralPath $OsppPath)) {
        return [PSCustomObject]@{ Success = $false; Output = ''; TimedOut = $false; VbsBlocked = $false }
    }

    try {
        $output = & cscript.exe //nologo $OsppPath @Arguments 2>&1 | Out-String
    } catch {
        return [PSCustomObject]@{
            Success = $false; Output = [string]$_.Exception.Message; TimedOut = $false; VbsBlocked = $false
        }
    }

    $output = $output.Trim()
    if (Test-OsppVbsBlockedOutput -Text $output) {
        return [PSCustomObject]@{ Success = $false; Output = $output; TimedOut = $false; VbsBlocked = $true }
    }
    if ($output -match 'Office \d+ Client Software License Management Tool') {
        return [PSCustomObject]@{ Success = $false; Output = ''; TimedOut = $false; VbsBlocked = $false }
    }

    return [PSCustomObject]@{ Success = $true; Output = $output; TimedOut = $false; VbsBlocked = $false }
}

function Normalize-ProductKey {
    param([string]$Key)
    return (($Key -replace '\s', '').ToUpper())
}

function Test-ProductKeyFormat {
    param([string]$Key)
    return (Normalize-ProductKey -Key $Key) -match '^[A-Z0-9]{5}(?:-[A-Z0-9]{5}){4}$'
}

function Get-OfficeLicenseInfo {
    param([Parameter(Mandatory)][string]$OsppPath)

    if (Test-OfficeLicenseUsesWmi) {
        return Get-OfficeLicenseInfoFromWmi
    }

    $result = Invoke-OsppScript -OsppPath $OsppPath -Arguments @('/dstatus')
    if ($result.VbsBlocked -or (Test-OsppVbsBlockedOutput -Text $result.Output)) {
        return Get-OfficeLicenseInfoFromWmi
    }
    if ($result.TimedOut) {
        return [PSCustomObject]@{
            Available = $false; TimedOut = $true; Licensed = $false
            LicenseStatus = $null; ProductKeyLast5 = $null; KmsHost = $null; Error = $null
        }
    }
    if (-not $result.Success -or [string]::IsNullOrWhiteSpace($result.Output)) {
        $wmiInfo = Get-OfficeLicenseInfoFromWmi
        if ($wmiInfo.ProductKeyLast5 -or $wmiInfo.LicenseStatus) { return $wmiInfo }
        return [PSCustomObject]@{
            Available = $false; TimedOut = $false; Licensed = $false
            LicenseStatus = $null; ProductKeyLast5 = $null; KmsHost = $null
            Error = 'Не вдалося отримати статус ліцензії (ospp.vbs)'
        }
    }

    $raw = $result.Output
    $info = [ordered]@{
        Available = $true; TimedOut = $false; Licensed = $false
        LicenseStatus = $null; LicenseDescription = $null; LicenseName = $null
        ProductKeyLast5 = $null; KmsHost = $null; Error = $null; Raw = $raw
    }

    if ($raw -match 'LICENSE STATUS:\s+([^\r\n]+)') {
        $info.LicenseStatus = $Matches[1].Trim()
        $info.Licensed = ($info.LicenseStatus -eq '---LICENSED---')
    }
    if ($raw -match 'LICENSE DESCRIPTION:\s+([^\r\n]+)') { $info.LicenseDescription = $Matches[1].Trim() }
    if ($raw -match 'Last 5 characters of installed product key:\s+(\S+)') { $info.ProductKeyLast5 = $Matches[1].Trim() }
    if ($raw -match 'KMS machine name:\s+([^\r\n]+)') { $info.KmsHost = $Matches[1].Trim() }
    if ($raw -match 'ERROR DESCRIPTION:\s+([^\r\n]+)') { $info.Error = $Matches[1].Trim() }
    $info.Backend = 'Ospp'

    return [PSCustomObject]$info
}

function Show-OfficeLicenseStatus {
    param([string]$OsppPath, [switch]$Detailed)

    if (-not $OsppPath) { $OsppPath = Get-OsppScriptPath }
    if (-not $OsppPath -and -not (Test-OfficeLicenseUsesWmi)) {
        Write-StatusLine '?' 'OSPP.VBS не знайдено' 'Yellow'
        return $null
    }
    if (-not $OsppPath) { $OsppPath = 'N/A' }

    $info = Get-OfficeLicenseInfo -OsppPath $OsppPath
    if ($info.TimedOut) {
        Write-StatusLine '?' 'Перевірка активації занадто довга' 'Yellow'
        return $info
    }
    if ($info.Error) { Write-StatusLine 'X' $info.Error 'Red' }
    elseif (-not $info.Available) {
        Write-StatusLine '?' 'Не вдалося визначити статус ліцензії' 'Yellow'
    } elseif ($info.Licensed) {
        Write-StatusLine 'OK' 'Office активовано' 'Green'
    } elseif ($info.LicenseStatus -eq '---UNLICENSED---') {
        Write-StatusLine '!' 'Office НЕ активовано' 'Yellow'
    } elseif ($info.LicenseStatus -eq '---NOTIFICATIONS---') {
        Write-StatusLine '!' 'Office потребує активації' 'Yellow'
    } elseif ($info.LicenseStatus) {
        Write-StatusLine '?' "Статус: $($info.LicenseStatus)" 'Yellow'
    } else {
        Write-StatusLine '?' 'Не вдалося визначити статус ліцензії' 'Yellow'
    }

    if ($Detailed) {
        if ($info.LicenseDescription) { Write-Host "    Опис:   $($info.LicenseDescription)" }
        if ($info.ProductKeyLast5) { Write-Host "    Ключ:   ...$($info.ProductKeyLast5)" }
        if ($info.KmsHost) { Write-Host "    KMS:    $($info.KmsHost)" }
        if ($info.Error) { Write-Host "    Помилка: $($info.Error)" -ForegroundColor Red }
    }
    return $info
}

function Invoke-OfficeOnlineActivation {
    param([Parameter(Mandatory)][string]$OsppPath)

    if (Test-OfficeLicenseUsesWmi) {
        return (Invoke-OfficeOnlineActivationWmi)
    }

    Write-Host ""
    Write-StatusLine 'i' 'Активація Office (ospp.vbs /act)...' 'Cyan'
    Write-Host "      Потрібен доступ до серверів Microsoft або KMS." -ForegroundColor Gray

    $act = Invoke-OsppScript -OsppPath $OsppPath -Arguments @('/act') -TimeoutSeconds 120
    if ($act.TimedOut) {
        Write-StatusLine '!' 'Активація занадто довга' 'Yellow'
        return 1
    }
    if (-not $act.Success) {
        if ($act.VbsBlocked -or (Test-OsppVbsBlockedOutput -Text $act.Output)) {
            return (Invoke-OfficeOnlineActivationWmi)
        }
        Write-StatusLine 'X' 'Не вдалося запустити ospp.vbs /act' 'Red'
        return 1
    }

    $info = Get-OfficeLicenseInfo -OsppPath $OsppPath
    if ($info.Licensed) {
        Write-StatusLine 'OK' 'Office активовано успішно' 'Green'
        return 0
    }
    if ($info.Error) { Write-StatusLine 'X' $info.Error 'Red' }
    else { Write-StatusLine '!' 'Office ще не активовано' 'Yellow' }
    return 1
}

function Invoke-OfficeProductKeyActivation {
    param(
        [Parameter(Mandatory)][string]$OsppPath,
        [Parameter(Mandatory)][string]$ProductKey
    )

    if (Test-OfficeLicenseUsesWmi) {
        return (Invoke-OfficeProductKeyActivationWmi -ProductKey $ProductKey)
    }

    $key = Normalize-ProductKey -Key $ProductKey
    if (-not (Test-ProductKeyFormat -Key $key)) {
        Write-StatusLine 'X' 'Невірний формат ключа' 'Red'
        return 2
    }

    Write-Host ""
    Write-StatusLine 'i' 'Встановлення ключа продукту...' 'Cyan'
    $install = Invoke-OsppScript -OsppPath $OsppPath -Arguments @("/inpkey:$key")
    if ($install.VbsBlocked -or (Test-OsppVbsBlockedOutput -Text $install.Output)) {
        return (Invoke-OfficeProductKeyActivationWmi -ProductKey $key)
    }
    Write-StatusLine 'OK' 'Ключ встановлено' 'Green'
    return (Invoke-OfficeOnlineActivation -OsppPath $OsppPath)
}

function Invoke-OfficeKmsActivation {
    param(
        [Parameter(Mandatory)][string]$OsppPath,
        [Parameter(Mandatory)][string]$KmsHost,
        [int]$KmsPort = 1688
    )

    if (Test-OfficeLicenseUsesWmi) {
        return (Invoke-OfficeKmsActivationWmi -KmsHost $KmsHost -KmsPort $KmsPort)
    }

    $hostName = $KmsHost.Trim()
    if ([string]::IsNullOrWhiteSpace($hostName)) {
        Write-StatusLine 'X' 'KMS-сервер не вказано' 'Red'
        return 2
    }
    if ($KmsPort -le 0) { $KmsPort = 1688 }

    Write-Host ""
    Write-StatusLine 'i' "Налаштування KMS: ${hostName}:${KmsPort}" 'Cyan'
    $setHost = Invoke-OsppScript -OsppPath $OsppPath -Arguments @("/sethst:$hostName")
    if ($setHost.VbsBlocked -or (Test-OsppVbsBlockedOutput -Text $setHost.Output)) {
        return (Invoke-OfficeKmsActivationWmi -KmsHost $hostName -KmsPort $KmsPort)
    }
    if ($KmsPort -ne 1688) {
        $null = Invoke-OsppScript -OsppPath $OsppPath -Arguments @("/setprt:$KmsPort")
    }

    $info = Get-OfficeLicenseInfo -OsppPath $OsppPath
    if (-not $info.ProductKeyLast5) {
        Write-Host "  [i] Спочатку введіть ключ Volume/MAK (пункт [1] у меню активації)." -ForegroundColor Yellow
        return 2
    }
    return (Invoke-OfficeOnlineActivation -OsppPath $OsppPath)
}

function Invoke-ActivateOffice {
    Initialize-UkrainianConsole
    $ErrorActionPreference = 'Continue'

    Write-Host ""
    Write-StatusLine 'i' 'Активація Microsoft Office' 'Cyan'
    Write-Host ""

    if (@(Get-InstalledOfficeInfo).Count -eq 0) {
        Write-StatusLine 'X' 'Office не встановлено — спочатку пункт [3]' 'Red'
        exit 1
    }

    $ospp = Get-OsppScriptPath
    $useWmi = Test-OfficeLicenseUsesWmi
    if (-not $useWmi -and -not $ospp) {
        Write-StatusLine 'X' 'OSPP.VBS не знайдено' 'Red'
        exit 1
    }
    if ($useWmi) {
        Write-Host "  [i] Виконання .vbs заблоковано — активація через PowerShell (WMI/CIM)" -ForegroundColor Yellow
        Write-Host ""
    }
    if (-not $ospp) { $ospp = 'N/A' }

    Write-Host "  Поточний статус:" -ForegroundColor Yellow
    $info = Show-OfficeLicenseStatus -OsppPath $ospp -Detailed

    if ($ProductKey) { exit (Invoke-OfficeProductKeyActivation -OsppPath $ospp -ProductKey $ProductKey) }
    if ($KmsHost) { exit (Invoke-OfficeKmsActivation -OsppPath $ospp -KmsHost $KmsHost -KmsPort $KmsPort) }
    if ($ActivateOnly) { exit (Invoke-OfficeOnlineActivation -OsppPath $ospp) }

    if ($info -and $info.Licensed) {
        Write-Host ""
        Write-StatusLine 'OK' 'Активація не потрібна' 'Green'
        exit 0
    }

    $exitCode = 0
    while ($true) {
        Write-Host ""
        Write-Host "  Оберіть дію:" -ForegroundColor Yellow
        Write-Host "    [1] Ввести ключ продукту (MAK / Retail)"
        Write-Host "    [2] Активація через KMS"
        Write-Host "    [3] Активувати зараз (ключ уже встановлено)"
        Write-Host "    [0] Скасувати"
        Write-Host ""
        $choice = Read-Host "  Ваш вибір (0-3)"

        switch ($choice) {
            '0' { exit 0 }
            '1' {
                Write-Host ""
                Write-Host "  Формат: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX" -ForegroundColor Gray
                $key = Read-Host "  Ключ продукту"
                if ([string]::IsNullOrWhiteSpace($key)) { continue }
                $exitCode = Invoke-OfficeProductKeyActivation -OsppPath $ospp -ProductKey $key
                break
            }
            '2' {
                Write-Host ""
                $kmsServer = Read-Host "  Адреса KMS-сервера"
                if ([string]::IsNullOrWhiteSpace($kmsServer)) { continue }
                $portInput = Read-Host "  Порт KMS [1688]"
                $port = 1688
                if (-not [string]::IsNullOrWhiteSpace($portInput)) {
                    if (-not [int]::TryParse($portInput, [ref]$port)) { continue }
                }
                $exitCode = Invoke-OfficeKmsActivation -OsppPath $ospp -KmsHost $kmsServer -KmsPort $port
                break
            }
            '3' {
                $exitCode = Invoke-OfficeOnlineActivation -OsppPath $ospp
                break
            }
            default { Write-Host "  [!] Невірний вибір" -ForegroundColor Yellow; continue }
        }
        if ($choice -in @('1', '2', '3')) { break }
    }

    Write-Host ""
    Write-Host "  Підсумок:" -ForegroundColor Yellow
    Show-OfficeLicenseStatus -OsppPath $ospp -Detailed | Out-Null
    exit $exitCode
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
    $null = Show-OfficeLicenseStatus -Detailed

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
        if ($officeEntry -and $installedApps.Count -ge 3) {
            Write-StatusLine "OK" "Office встановлено (setup.exe: код $SetupExitCode, у логах: успіх)" "Green"
            exit 0
        }
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

function Test-OfficeSetupLogLineIsRemoval {
    param(
        [string]$Line,
        [ValidateSet('Install', 'Uninstall', 'RemoveApps', 'Unknown')]
        [string]$SetupOperation = 'Unknown'
    )

    if ($SetupOperation -in @('Uninstall', 'RemoveApps')) { return $true }

    if ($Line -match '(?i)configuration-remove|ODTUninstall|UNINSTALLCENTENNIAL|ProductsToRemove|"Scenario"\s*:\s*"REMOVE"|ScenarioSubType.*Uninstall|TaskType.*UNINSTALL|ProductsToRemove') {
        return $true
    }

    return $false
}

function Convert-OfficeSetupLogLine {
    param(
        [string]$Line,
        [ValidateSet('Install', 'Uninstall', 'RemoveApps', 'Unknown')]
        [string]$SetupOperation = 'Unknown'
    )

    if ([string]::IsNullOrWhiteSpace($Line)) { return $null }

    $line = $Line.Trim()
    if ($line.Length -lt 8) { return $null }

    $isRemoval = Test-OfficeSetupLogLineIsRemoval -Line $line -SetupOperation $SetupOperation

    if ($line -match '(?i)Telemetry Event|ActivityEnded|SendEvent|DroppedAggregatedActivity|General Telemetry') {
        return $null
    }

    if ($line -match '(?i)(Non Task Error|Task Error)\s+\S+\s+Unexpected\s+') {
        $text = 'Помилка Office Click-to-Run'
        if ($line -match '"ContextData"\s*:\s*"([^"]+)"') {
            $text = $Matches[1]
        } elseif ($line -match '"ErrorMessage"\s*:\s*"([^"]+)"') {
            $text = $Matches[1]
        }
        if ($line -match '"ErrorCode"\s*:\s*(-?\d+)') {
            $code = $Matches[1]
            if ($code -ne '0') {
                $text = if ($text -eq 'Помилка Office Click-to-Run') { "Код помилки Office: $code" } else { "[$code] $text" }
            }
        }
        return [PSCustomObject]@{ Kind = 'Error'; Text = $text }
    }

    if ($line -match 'total progress is now (\d+)\.') {
        $pct = [int]$Matches[1]
        if ($pct -ge 0 -and $pct -le 100) {
            $label = if ($isRemoval) { 'Видалення' } else { 'Прогрес' }
            return [PSCustomObject]@{
                Kind    = 'Percent'
                Text    = "$label`: $pct%"
                Percent = $pct
            }
        }
    }

    if ($line -match 'splash progress is now (\d+)\.') {
        $pct = [int]$Matches[1]
        if ($pct -ge 0 -and $pct -le 120) {
            $norm = [Math]::Min(100, [int][Math]::Round($pct * 100.0 / 120.0))
            $label = if ($isRemoval) { 'Видалення' } else { 'Прогрес' }
            return [PSCustomObject]@{
                Kind    = 'Percent'
                Text    = "$label`: $norm%"
                Percent = $norm
            }
        }
    }

    if ($line -match '(?i)Exit with error code (\d+)') {
        return [PSCustomObject]@{ Kind = 'Error'; Text = "Office повернув код $($Matches[1])" }
    }

    if ($line -match '\s(Medium|Verbose|Monitorable)\s+') {
        if ($line -match '"PercentComplete"\s*:\s*(\d{1,3})') {
            $pct = [int]$Matches[1]
            if ($pct -ge 0 -and $pct -le 100) {
                $label = if ($isRemoval) { 'Видалення' } else { 'Прогрес' }
                return [PSCustomObject]@{
                    Kind    = 'Percent'
                    Text    = "$label`: $pct%"
                    Percent = $pct
                }
            }
        }

        if ($line -match '(?i)"message"\s*:\s*"(Downloading|Installing|Applying|Copying|Removing|Updating)[^"]*"') {
            $stage = $Matches[1]
            if ($isRemoval -and $stage -match '(?i)^Installing$') { $stage = 'Removing' }
            $uk = switch -Regex ($stage) {
                '^Downloading$' { 'Завантаження' }
                '^Installing$'  { if ($isRemoval) { 'Видалення' } else { 'Встановлення' } }
                '^Applying$'    { 'Застосування' }
                '^Copying$'     { 'Копіювання' }
                '^Removing$'    { 'Видалення' }
                '^Updating$'    { 'Оновлення' }
                default         { $stage }
            }
            return [PSCustomObject]@{ Kind = 'Stage'; Text = "$uk..." }
        }

        if ($line -match '(?i)C2R::Setup::') {
            if ($line -match '(?i)remove|uninstall') {
                return [PSCustomObject]@{ Kind = 'Stage'; Text = 'Видалення Office...' }
            }
            if ($line -match '(?i)download') {
                return [PSCustomObject]@{ Kind = 'Stage'; Text = 'Підготовка компонентів...' }
            }
            if ($line -match '(?i)\bconfigure\b|\binstall\b') {
                $text = if ($isRemoval) { 'Видалення компонентів...' } else { 'Встановлення компонентів...' }
                return [PSCustomObject]@{ Kind = 'Stage'; Text = $text }
            }
            if ($line -match '(?i)register') {
                return [PSCustomObject]@{ Kind = 'Stage'; Text = 'Реєстрація Office...' }
            }
        }

        foreach ($app in @('Word', 'Excel', 'Outlook', 'PowerPoint', 'Access', 'Publisher', 'OneNote')) {
            if ($line -match "\b$app\b" -and $line -match '(?i)package|stream|app|remove|uninstall') {
                $verb = if ($isRemoval) { 'Видалення' } else { 'Встановлення' }
                return [PSCustomObject]@{ Kind = 'App'; Text = "$verb`: $app" }
            }
        }

        if ($line -match '(?i)"Scenario"\s*:\s*"REMOVE"|ODTUninstall|UNINSTALLCENTENNIAL|ProductsToRemove') {
            return [PSCustomObject]@{ Kind = 'Stage'; Text = 'Видалення Office...' }
        }

        if ($line -match '(?i)Bootstrapper Finished') {
            return [PSCustomObject]@{ Kind = 'Stage'; Text = 'Завершення...' }
        }
    }

    return $null
}

function Get-OfficeSetupLogPriority {
    param([string]$Name)

    if ($Name -match '^(?i)PPO-') { return 300 }
    if ($Name -match '^(?i)DESKTOP-') { return 200 }
    if ($Name -match '^(?i)OfficeSetup_') { return 150 }
    if ($Name -match '(?i)officeclicktorun|aria-debug') { return 10 }
    if ($Name -match '\.log$') { return 50 }

    return 0
}

function Get-OfficeSetupLogCandidates {
    param(
        [string[]]$SearchPaths,
        [DateTime]$NotBefore
    )

    $files = @()
    foreach ($dir in $SearchPaths) {
        if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { continue }

        $files += Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue |
            Where-Object {
                $_.LastWriteTime -ge $NotBefore.AddMinutes(-1) -and
                ($_.Extension -eq '.log' -or $_.Name -match '(?i)setup|office|c2r|desktop|ppo')
            }
    }

    return @(
        $files |
            Sort-Object @{ Expression = { Get-OfficeSetupLogPriority -Name $_.Name }; Descending = $true },
                      LastWriteTime -Descending
    )
}

function Get-ActiveOfficeSetupLogFile {
    param(
        [string[]]$SearchPaths,
        [DateTime]$NotBefore
    )

    return Get-OfficeSetupLogCandidates -SearchPaths $SearchPaths -NotBefore $NotBefore |
        Select-Object -First 1
}

function Save-OfficeSetupLogCopy {
    param(
        [System.IO.FileInfo]$SourceLog,
        [string]$LogDirectory
    )

    if (-not $SourceLog -or -not (Test-Path -LiteralPath $SourceLog.FullName)) { return $null }
    if (-not $LogDirectory) { return $SourceLog.FullName }

    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }

    $dest = Join-Path $LogDirectory ("OfficeSetup_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
    Copy-Item -LiteralPath $SourceLog.FullName -Destination $dest -Force
    return $dest
}

function Get-OfficeSetupExitCodeFromLog {
    param([string]$LogPath)

    if (-not $LogPath -or -not (Test-Path -LiteralPath $LogPath)) { return $null }

    try {
        $lines = Get-Content -LiteralPath $LogPath -Tail 300 -ErrorAction Stop
    } catch {
        return $null
    }

    for ($i = $lines.Count - 1; $i -ge 0; $i--) {
        $line = $lines[$i]
        if ($line -match 'Bootstrapper Finished.*?ExitCode\\":\\"(\d+)\\"') {
            return [int]$Matches[1]
        }
        if ($line -match 'Bootstrapper Finished.*?"ExitCode"\s*:\s*"(\d+)"') {
            return [int]$Matches[1]
        }
        if ($line -match '(?i)UniversalBootstrapper\.Application3.*?"Data\.ExitCode"\s*:\s*(\d+)') {
            return [int]$Matches[1]
        }
        if ($line -match '(?i)Main:\s+returned:\s+(\d+)') {
            return [int]$Matches[1]
        }
        if ($line -match '(?i)Exit with error code\s+(\d+)') {
            return [int]$Matches[1]
        }
    }

    for ($i = $lines.Count - 1; $i -ge 0; $i--) {
        $line = $lines[$i]
        if ($line -match 'ExitCode\\":\\"(\d+)\\"') {
            return [int]$Matches[1]
        }
        if ($line -match '"ExitCode"\s*:\s*"(\d+)"') {
            return [int]$Matches[1]
        }
        if ($line -match '"ErrorCode"\s*:\s*(-?\d+)') {
            $code = [int]$Matches[1]
            if ($code -ne 0) { return $code }
        }
    }

    return $null
}

function Get-OfficeSetupExitCodeFromLogs {
    param(
        [string[]]$SearchPaths,
        [DateTime]$NotBefore
    )

    foreach ($file in (Get-OfficeSetupLogCandidates -SearchPaths $SearchPaths -NotBefore $NotBefore)) {
        $code = Get-OfficeSetupExitCodeFromLog -LogPath $file.FullName
        if ($null -ne $code) { return $code }
    }

    return $null
}

function Get-OfficeSetupErrorsFromLog {
    param([string]$LogPath)

    if (-not $LogPath -or -not (Test-Path -LiteralPath $LogPath)) { return @() }

    $errors = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($line in (Get-Content -LiteralPath $LogPath -ErrorAction Stop)) {
            $parsed = Convert-OfficeSetupLogLine -Line $line
            if ($parsed -and $parsed.Kind -eq 'Error') {
                if (-not $errors.Contains($parsed.Text)) {
                    $errors.Add($parsed.Text)
                }
            }
        }
    } catch { }

    return @($errors)
}

function Get-OfficeInstallRelatedProcesses {
    @('setup', 'OfficeClickToRun', 'OfficeC2RClient', 'OfficeSetup') | ForEach-Object {
        Get-Process -Name $_ -ErrorAction SilentlyContinue
    }
}

function Test-OfficeAppsInstalled {
    $roots = @(
        'C:\Program Files\Microsoft Office\root\Office16',
        'C:\Program Files\Microsoft Office\Office16'
    )
    $apps = @('WINWORD.EXE', 'EXCEL.EXE', 'OUTLOOK.EXE')

    foreach ($root in $roots) {
        foreach ($app in $apps) {
            if (Test-Path -LiteralPath (Join-Path $root $app)) { return $true }
        }
    }

    return $false
}

function Test-PartialOfficeClickToRunInstall {
    $ctr = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
    if (-not $ctr) { return $false }

    $clientPath = 'C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe'
    if (-not (Test-Path -LiteralPath $clientPath)) { return $false }

    return -not (Test-OfficeAppsInstalled)
}

function Stop-OfficeSetupProcesses {
    Get-Process -Name setup -ErrorAction SilentlyContinue | ForEach-Object {
        try { Stop-Process -Id $_.Id -Force -ErrorAction Stop } catch { }
    }
}

function Stop-OfficeClickToRunEnvironment {
    param([switch]$Quiet)

    if (-not $Quiet) {
        Write-Host "  [i] Зупинка процесів Office Click-to-Run..." -ForegroundColor Cyan
    }

    Stop-OfficeSetupProcesses

    Get-OfficeInstallRelatedProcesses | ForEach-Object {
        try { Stop-Process -Id $_.Id -Force -ErrorAction Stop } catch { }
    }

    $svc = Get-Service -Name ClickToRunSvc -ErrorAction SilentlyContinue
    if ($svc -and $svc.Status -ne 'Stopped') {
        try {
            Stop-Service -Name ClickToRunSvc -Force -ErrorAction Stop
            $svc.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Stopped, (New-TimeSpan -Seconds 45))
        } catch { }
    }

    Start-Sleep -Seconds 2

    Get-OfficeInstallRelatedProcesses | ForEach-Object {
        try { Stop-Process -Id $_.Id -Force -ErrorAction Stop } catch { }
    }
}

function Start-OfficeClickToRunService {
    param([switch]$Quiet)

    $svc = Get-Service -Name ClickToRunSvc -ErrorAction SilentlyContinue
    if (-not $svc) { return $false }

    if ($svc.Status -ne 'Running') {
        if (-not $Quiet) {
            Write-Host "  [i] Запуск служби ClickToRunSvc..." -ForegroundColor Cyan
        }
        try {
            Start-Service -Name ClickToRunSvc -ErrorAction Stop
            $svc.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Running, (New-TimeSpan -Seconds 60))
        } catch { }
    }

    return ($svc.Status -eq 'Running')
}

function Invoke-CleanupPartialOfficeInstall {
    param([string]$InstallDir)

    $setupPath = Join-Path $InstallDir 'setup.exe'
    if (-not (Test-Path -LiteralPath $setupPath)) { return 1 }

    Stop-OfficeSetupProcesses
    Start-OfficeClickToRunService | Out-Null

    Invoke-PrepareOfficeConfig -InstallDir $InstallDir -Mode Uninstall | Out-Null
    $configPath = Join-Path $InstallDir 'configuration-remove.xml'

    Write-Host "  [i] Видалення залишків незавершеного встановлення..." -ForegroundColor Yellow
    Write-Host "      (тихо, без вікна Microsoft)" -ForegroundColor DarkGray

    $logDir = Get-OfficeInstallerLogDirectory -InstallDir $InstallDir
    $startedAt = Get-Date
    $envState = Use-OfficeSetupLogDirectory -LogDirectory $logDir

    try {
        $proc = Start-Process -FilePath $setupPath -ArgumentList @('/configure', $configPath) -PassThru -NoNewWindow
        $exitCode = Wait-OfficeSetupWithProgress `
            -Process $proc `
            -LogDirectory $logDir `
            -StartedAt $startedAt `
            -ShowProgress:$false `
            -SetupOperation Uninstall
    } finally {
        Restore-OfficeSetupEnvironment -State $envState
        if ($logDir) {
            Sync-OfficeSetupLogsToDirectory -LogDirectory $logDir -NotBefore $startedAt | Out-Null
        }
    }

    Stop-OfficeClickToRunEnvironment -Quiet

    return $exitCode
}

function Prepare-OfficeInstallEnvironment {
    param(
        [string]$InstallDir,
        [switch]$Reinstall
    )

    Write-Host ""
    Write-Host "  [i] Підготовка середовища для встановлення..." -ForegroundColor Cyan

    Stop-OfficeSetupProcesses

    if ($Reinstall) {
        Write-Host "  [OK] Готово до перевстановлення (видалення — у конфігурації install)" -ForegroundColor Green
        return
    }

    if (Test-PartialOfficeClickToRunInstall) {
        Write-Host "  [!] Знайдено залишки попередньої спроби (C2R без програм Office)" -ForegroundColor Yellow

        $removeCode = Invoke-CleanupPartialOfficeInstall -InstallDir $InstallDir
        if ($removeCode -ne 0 -and $removeCode -ne 3010) {
            Write-Host "  [!] Видалення залишків повернуло код $removeCode — продовжуємо..." -ForegroundColor Yellow
        } else {
            Write-Host "  [OK] Залишки видалено" -ForegroundColor Green
        }

        Start-Sleep -Seconds 3
    } else {
        Write-Host "  [OK] Система готова до чистого встановлення" -ForegroundColor Green
    }
}

function Get-OfficeSetupErrorHelp {
    param([int]$ExitCode)

    switch ($ExitCode) {
        17004 {
            return @(
                'Office Click-to-Run не зміг завершити конфігурацію (код 17004).',
                'Типова причина: залишки попередньої спроби або служба ClickToRunSvc не відповіла.',
                'Спробуйте: перезавантажити ПК, потім пункт [4] «Перевстановити» або [5] «Видалити Office».'
            ) -join ' '
        }
        30182 { return 'Office не знайшов або не зміг прочитати офлайн-файли. Перевірте, що папка Office\Data містить ~2 GB файлів stream.x64.*.dat.' }
        30053 { return 'Служба ClickToRunSvc була недоступна під час підготовки. Якщо основне встановлення завершилось успішно — це можна ігнорувати.' }
        0     { return '' }
        3010  { return '' }
        default { return 'Деталі у логах у папці Logs (файли PPO-*.log).' }
    }
}

function Get-OfficeInstallerLogDirectory {
    param([string]$InstallDir)

    if (-not $InstallDir) { return $null }

    $logDir = Join-Path $InstallDir 'Logs'
    if (-not (Test-Path -LiteralPath $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    return (Resolve-Path -LiteralPath $logDir).Path
}

function Get-OfficeSetupTempSearchPaths {
    param([string[]]$ExtraPaths)

    $dirs = New-Object System.Collections.Generic.List[string]

    foreach ($path in @($env:TEMP, $env:TMP, $(Join-Path $env:LOCALAPPDATA 'Temp'), $(Join-Path $env:WINDIR 'Temp'))) {
        if (-not $path) { continue }
        try {
            if (-not (Test-Path -LiteralPath $path)) { continue }
            $resolved = (Resolve-Path -LiteralPath $path).Path
            if (-not $dirs.Contains($resolved)) { [void]$dirs.Add($resolved) }
        } catch { }
    }

    foreach ($path in @($ExtraPaths)) {
        if (-not $path) { continue }
        try {
            if (-not (Test-Path -LiteralPath $path)) { continue }
            $resolved = (Resolve-Path -LiteralPath $path).Path
            if (-not $dirs.Contains($resolved)) { [void]$dirs.Add($resolved) }
        } catch { }
    }

    return @($dirs)
}

function Sync-OfficeSetupLogsToDirectory {
    param(
        [string]$LogDirectory,
        [DateTime]$NotBefore
    )

    if (-not $LogDirectory) { return @() }

    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }

    $patterns = @('PPO-*.log', 'DESKTOP-*.log', 'officeclicktorun*.log', 'aria-debug*.log')
    $synced = New-Object System.Collections.Generic.List[string]
    $cutoff = $NotBefore.AddMinutes(-2)

    $logResolved = (Resolve-Path -LiteralPath $LogDirectory).Path

    foreach ($sourceDir in (Get-OfficeSetupTempSearchPaths -ExtraPaths @($LogDirectory))) {
        if ($sourceDir -ieq $logResolved) { continue }

        foreach ($pattern in $patterns) {
            $files = Get-ChildItem -LiteralPath $sourceDir -Filter $pattern -File -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -ge $cutoff }

            foreach ($file in $files) {
                $dest = Join-Path $LogDirectory $file.Name
                $needsCopy = $true

                if (Test-Path -LiteralPath $dest) {
                    $destInfo = Get-Item -LiteralPath $dest
                    if ($destInfo.Length -ge $file.Length -and $destInfo.LastWriteTime -ge $file.LastWriteTime) {
                        $needsCopy = $false
                    }
                }

                if ($needsCopy) {
                    try {
                        Copy-Item -LiteralPath $file.FullName -Destination $dest -Force
                    } catch { }
                }

                if ((Test-Path -LiteralPath $dest) -and -not $synced.Contains($dest)) {
                    [void]$synced.Add($dest)
                }
            }
        }
    }

    Get-ChildItem -LiteralPath $LogDirectory -Filter '*.log' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $cutoff } |
        ForEach-Object {
            if (-not $synced.Contains($_.FullName)) { [void]$synced.Add($_.FullName) }
        }

    return @($synced | Sort-Object)
}

function Write-OfficeSetupLogSummary {
    param(
        [string]$LogDirectory,
        [DateTime]$NotBefore
    )

    if (-not $LogDirectory -or -not (Test-Path -LiteralPath $LogDirectory)) { return }

    $files = @(Get-ChildItem -LiteralPath $LogDirectory -Filter '*.log' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $NotBefore.AddMinutes(-2) } |
        Sort-Object LastWriteTime -Descending)

    if ($files.Count -eq 0) { return }

    Write-Host "  [i] Логи Office: $LogDirectory" -ForegroundColor DarkGray
    foreach ($file in ($files | Select-Object -First 4)) {
        Write-Host "      $($file.Name)" -ForegroundColor DarkGray
    }
    if ($files.Count -gt 4) {
        Write-Host "      ... і ще $($files.Count - 4) файл(ів)" -ForegroundColor DarkGray
    }
}

function Get-OfficeSetupWorkDirectory {
    param([string]$LogDirectory)

    if (-not $LogDirectory) { return $null }

    $workDir = Join-Path $LogDirectory '_odt_work'
    if (-not (Test-Path -LiteralPath $workDir)) {
        New-Item -ItemType Directory -Path $workDir -Force | Out-Null
    }

    return (Resolve-Path -LiteralPath $workDir).Path
}

function Use-OfficeSetupLogDirectory {
    param([string]$LogDirectory)

    $state = [ordered]@{
        Temp = $env:TEMP
        Tmp  = $env:TMP
    }

    if ($LogDirectory) {
        $workDir = Get-OfficeSetupWorkDirectory -LogDirectory $LogDirectory
        if ($workDir) {
            $env:TEMP = $workDir
            $env:TMP = $workDir
        }
    }

    return $state
}

function Restore-OfficeSetupEnvironment {
    param($State)

    if (-not $State) { return }
    $env:TEMP = $State.Temp
    $env:TMP = $State.Tmp
}

function Resolve-ProcessExitCode {
    param(
        [System.Diagnostics.Process]$Process,
        [string]$LogPath,
        [string[]]$SearchPaths,
        [DateTime]$NotBefore
    )

    if ($Process) {
        try {
            if (-not $Process.HasExited) {
                $Process.WaitForExit()
            }
            $Process.Refresh()
        } catch { }
    }

    $fromLog = $null
    if ($SearchPaths -and $NotBefore) {
        $fromLog = Get-OfficeSetupExitCodeFromLogs -SearchPaths $SearchPaths -NotBefore $NotBefore
    }
    if ($null -eq $fromLog -and $LogPath) {
        $fromLog = Get-OfficeSetupExitCodeFromLog -LogPath $LogPath
    }

    if ($null -ne $fromLog) { return [int]$fromLog }

    if ($Process) {
        try {
            $Process.Refresh()
            if ($null -ne $Process.ExitCode) {
                return [int]$Process.ExitCode
            }
        } catch { }
    }

    return 1
}

function Show-OfficeSetupProgressLine {
    param(
        [string]$Status,
        [int]$Percent = -1,
        [TimeSpan]$Elapsed
    )

    $barLen = 28
    if ($Percent -ge 0 -and $Percent -le 100) {
        $fill = [int][Math]::Round($barLen * $Percent / 100.0)
        $bar = ('#' * $fill) + ('-' * ($barLen - $fill))
        $pctStr = ('{0,3}%' -f $Percent)
    } else {
        $bar = ('~' * ([int](($Elapsed.TotalSeconds % 4)) + 1)).PadRight($barLen, '.')
        $pctStr = '   '
    }

    $statusShort = if ($Status.Length -gt 42) { $Status.Substring(0, 39) + '...' } else { $Status }
    $line = "  [$bar] $pctStr  $statusShort  ($($Elapsed.ToString('mm\:ss')))"
    Write-Host ("`r{0,-110}" -f $line) -NoNewline
}

function Get-ActiveOfficeSetupLogFile {
    param(
        [string[]]$SearchPaths,
        [DateTime]$NotBefore
    )

    $files = @()
    foreach ($dir in $SearchPaths) {
        if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { continue }
        $files += Get-ChildItem -LiteralPath $dir -Filter '*.log' -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -ge $NotBefore.AddSeconds(-5) }
    }

    return $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

function Wait-OfficeSetupWithProgress {
    param(
        [System.Diagnostics.Process]$Process,
        [string]$LogDirectory,
        [DateTime]$StartedAt,
        [int]$StallWarningSeconds = 180,
        [switch]$ShowProgress,
        [ValidateSet('Install', 'Uninstall', 'RemoveApps', 'Unknown')]
        [string]$SetupOperation = 'Unknown'
    )

    if (-not $Process) { return -1 }

    if (-not $PSBoundParameters.ContainsKey('ShowProgress')) {
        $ShowProgress = $true
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $lastStatus = if ($SetupOperation -eq 'Uninstall') { 'Запуск видалення...' } else { 'Запуск setup.exe...' }
    $lastPercent = -1
    $seenPos = 0
    $activeLog = $null
    $lastActivity = [DateTime]::UtcNow
    $warnedStall = $false
    $lastSync = [DateTime]::MinValue
    $workDir = if ($LogDirectory) { Get-OfficeSetupWorkDirectory -LogDirectory $LogDirectory } else { $null }
    $tempSearchPaths = Get-OfficeSetupTempSearchPaths -ExtraPaths @($LogDirectory, $workDir)
    $searchPaths = @()
    if ($LogDirectory) { $searchPaths += $LogDirectory }
    if ($workDir) { $searchPaths += $workDir }
    $searchPaths += $tempSearchPaths

    if ($ShowProgress) {
        Show-OfficeSetupProgressLine -Status $lastStatus -Percent $lastPercent -Elapsed $sw.Elapsed
    }

    while (-not $Process.HasExited) {
        Start-Sleep -Milliseconds 500

        if ($LogDirectory -and ((Get-Date) - $lastSync).TotalSeconds -ge 2) {
            Sync-OfficeSetupLogsToDirectory -LogDirectory $LogDirectory -NotBefore $StartedAt | Out-Null
            $lastSync = Get-Date
        }

        if (-not $activeLog) {
            $activeLog = Get-ActiveOfficeSetupLogFile -SearchPaths $searchPaths -NotBefore $StartedAt
            if ($activeLog) {
                $seenPos = 0
                $lastActivity = [DateTime]::UtcNow
                $lastStatus = "Читання логу: $($activeLog.Name)"
            }
        }

        if ($ShowProgress -and $activeLog -and (Test-Path -LiteralPath $activeLog.FullName)) {
            try {
                $stream = [System.IO.File]::Open($activeLog.FullName, [System.IO.FileMode]::Open, `
                    [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                try {
                    if ($stream.Length -gt $seenPos) {
                        $stream.Seek($seenPos, [System.IO.SeekOrigin]::Begin) | Out-Null
                        $reader = New-Object System.IO.StreamReader($stream)
                        while (-not $reader.EndOfStream) {
                            $line = $reader.ReadLine()
                            $parsed = Convert-OfficeSetupLogLine -Line $line -SetupOperation $SetupOperation
                            if (-not $parsed) { continue }

                            $lastActivity = [DateTime]::UtcNow
                            if ($parsed.Kind -eq 'Error') {
                                Write-Host ""
                                Write-Host "  [!] $($parsed.Text)" -ForegroundColor Red
                            } elseif ($parsed.Kind -eq 'Percent') {
                                $lastPercent = $parsed.Percent
                                $lastStatus = $parsed.Text
                            } else {
                                $lastStatus = $parsed.Text
                            }
                        }
                        $seenPos = $stream.Length
                    }
                } finally {
                    $stream.Dispose()
                }
            } catch { }
        } elseif ($ShowProgress -and $sw.Elapsed.TotalSeconds -ge 10) {
            $lastStatus = 'setup.exe працює, очікування логу...'
        }

        if ($ShowProgress) {
            $stallSec = ([DateTime]::UtcNow - $lastActivity).TotalSeconds
            if ($stallSec -ge $StallWarningSeconds -and -not $warnedStall) {
                $warnedStall = $true
                Write-Host ""
                Write-Host "  [!] Немає нових записів у логу $([int]$stallSec) сек — setup.exe ще працює, зачекайте..." -ForegroundColor Yellow
                if ($activeLog) {
                    Write-Host "      Лог: $($activeLog.FullName)" -ForegroundColor DarkGray
                } elseif ($LogDirectory) {
                    Write-Host "      Логи копіюються у: $LogDirectory" -ForegroundColor DarkGray
                } else {
                    Write-Host "      Шукали лог у: $($searchPaths -join ', ')" -ForegroundColor DarkGray
                }
            }

            Show-OfficeSetupProgressLine -Status $lastStatus -Percent $lastPercent -Elapsed $sw.Elapsed
        }
    }

    $postExitDeadline = [DateTime]::UtcNow.AddSeconds(120)
    while ([DateTime]::UtcNow -lt $postExitDeadline) {
        if ($LogDirectory -and ((Get-Date) - $lastSync).TotalSeconds -ge 2) {
            Sync-OfficeSetupLogsToDirectory -LogDirectory $LogDirectory -NotBefore $StartedAt | Out-Null
            $lastSync = Get-Date
        }

        $exitFromLog = Get-OfficeSetupExitCodeFromLogs -SearchPaths $searchPaths -NotBefore $StartedAt
        if ($null -ne $exitFromLog) { break }

        $stillRunning = @(Get-OfficeInstallRelatedProcesses | Where-Object { $_.ProcessName -ne 'setup' })
        if ($stillRunning.Count -eq 0) {
            Start-Sleep -Seconds 2
            break
        }

        if ($ShowProgress) {
            $lastStatus = 'Завершення Office Click-to-Run...'
            Show-OfficeSetupProgressLine -Status $lastStatus -Percent $lastPercent -Elapsed $sw.Elapsed
        }

        Start-Sleep -Milliseconds 500
    }

    if ($LogDirectory) {
        Sync-OfficeSetupLogsToDirectory -LogDirectory $LogDirectory -NotBefore $StartedAt | Out-Null
    }

    if (-not $activeLog) {
        $activeLog = Get-ActiveOfficeSetupLogFile -SearchPaths $searchPaths -NotBefore $StartedAt
    }

    $syncedLogs = if ($LogDirectory) {
        @(Sync-OfficeSetupLogsToDirectory -LogDirectory $LogDirectory -NotBefore $StartedAt)
    } else {
        @()
    }

    $primaryLog = Get-ActiveOfficeSetupLogFile -SearchPaths $searchPaths -NotBefore $StartedAt
    $logForExit = if ($primaryLog) { $primaryLog.FullName } elseif ($syncedLogs.Count -gt 0) { $syncedLogs[0] } else { $null }

    if (-not $logForExit -and $activeLog -and $LogDirectory) {
        $logForExit = Save-OfficeSetupLogCopy -SourceLog $activeLog -LogDirectory $LogDirectory
    } elseif (-not $logForExit -and $activeLog) {
        $logForExit = $activeLog.FullName
    }

    if ($ShowProgress -and $logForExit -and (Test-Path -LiteralPath $logForExit)) {
        try {
            $tail = Get-Content -LiteralPath $logForExit -Tail 30 -ErrorAction Stop
            foreach ($line in $tail) {
                $parsed = Convert-OfficeSetupLogLine -Line $line -SetupOperation $SetupOperation
                if ($parsed -and $parsed.Kind -eq 'Error') {
                    Write-Host "  [!] $($parsed.Text)" -ForegroundColor Red
                }
            }
        } catch { }
    }

    if ($LogDirectory) {
        Write-OfficeSetupLogSummary -LogDirectory $LogDirectory -NotBefore $StartedAt
    }

    if ($ShowProgress) {
        Write-Host ""
    }

    return (Resolve-ProcessExitCode -Process $Process -LogPath $logForExit -SearchPaths $searchPaths -NotBefore $StartedAt)
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
        [string]$Description,
        [string]$InstallDir,
        [switch]$ShowTerminalProgress,
        [ValidateSet('Install', 'Uninstall', 'RemoveApps', 'Unknown')]
        [string]$SetupOperation = 'Install'
    )

    Write-Host "  [i] $Description" -ForegroundColor Cyan

    $logDir = if ($InstallDir) { Get-OfficeInstallerLogDirectory -InstallDir $InstallDir } else { $null }

    if ($ShowTerminalProgress) {
        Write-Host "  [i] Тихий режим — прогрес відображається тут, у терміналі" -ForegroundColor Gray
    } else {
        Write-Host "      Це може зайняти кілька хвилин..." -ForegroundColor Gray
    }

    if ($logDir) {
        Write-Host "  [i] Логи Office: $logDir" -ForegroundColor DarkGray
    }

    $startedAt = Get-Date
    $envState = Use-OfficeSetupLogDirectory -LogDirectory $logDir

    try {
        $proc = Start-Process -FilePath $SetupPath -ArgumentList @('/configure', $ConfigPath) -PassThru -NoNewWindow
        $exitCode = Wait-OfficeSetupWithProgress `
            -Process $proc `
            -LogDirectory $logDir `
            -StartedAt $startedAt `
            -ShowProgress:$ShowTerminalProgress `
            -SetupOperation $SetupOperation
    } finally {
        Restore-OfficeSetupEnvironment -State $envState
        if ($logDir) {
            Sync-OfficeSetupLogsToDirectory -LogDirectory $logDir -NotBefore $startedAt | Out-Null
        }
    }

    if ($exitCode -eq 0 -or $exitCode -eq 3010) {
        Write-Host "  [OK] setup.exe завершено (код $exitCode)" -ForegroundColor Green
    } else {
        Write-Host "  [!] setup.exe повернув код $exitCode" -ForegroundColor Yellow
        if ($logDir) {
            Write-Host "      Деталі: $logDir" -ForegroundColor Yellow
        }
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
            -ConfigPath $config -InstallDir $InstallDir -ShowTerminalProgress `
            -SetupOperation RemoveApps `
            -Description 'Видалення Skype for Business з пакету Office'
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
    'Uninstall' {
        Invoke-UninstallOffice
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
    'Activate' {
        Invoke-ActivateOffice
    }
}
