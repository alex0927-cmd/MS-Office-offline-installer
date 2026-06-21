@echo off
chcp 1251 >nul 2>&1
setlocal EnableDelayedExpansion

set "ACTION="
if /I "%~1"=="/install" set "ACTION=install"
if /I "%~1"=="/reinstall" set "ACTION=reinstall"
if /I "%~1"=="/force" set "ACTION=reinstall"
if /I "%~1"=="/uninstall" set "ACTION=uninstall"
if /I "%~1"=="/download" set "ACTION=download"
if /I "%~1"=="/downloadodt" set "ACTION=downloadodt"
if /I "%~1"=="/removeteamsskype" set "ACTION=removeteamsskype"

net session >nul 2>&1
if %errorLevel% neq 0 (
    echo.
    echo  [УВАГА] Потрібні права адміністратора. Підтвердіть UAC...
    echo.
    if defined ACTION (
        powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -ArgumentList '%~1' -Verb RunAs"
    ) else (
        powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    )
    echo.
    echo  Відкрито окреме вікно з правами адміністратора.
    echo  Дочекайтеся завершення у тому вікні.
    timeout /t 3 >nul
    exit /b
)

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "ROOT_DIR=%SCRIPT_DIR%"
cd /d "%SCRIPT_DIR%"

:main_loop
set "START_TIME=%TIME%"
set "SKIP_INSTALL=0"
set "REINSTALL=0"
set "EXIT_CODE=0"
set "VERIFY_CODE=0"

if not defined ACTION goto :menu

if /I "%ACTION%"=="install" set "ACTION=" & goto :action_install
if /I "%ACTION%"=="reinstall" set "ACTION=" & goto :action_reinstall
if /I "%ACTION%"=="uninstall" set "ACTION=" & goto :action_uninstall
if /I "%ACTION%"=="download" set "ACTION=" & goto :action_download
if /I "%ACTION%"=="downloadodt" set "ACTION=" & goto :action_download_odt
if /I "%ACTION%"=="removeteamsskype" set "ACTION=" & goto :action_remove_teams_skype
goto :menu

:header
cls
echo.
echo  ============================================================
echo    Office 2021 Professional Plus - офлайн інсталятор
echo  ============================================================
echo.
echo  Папка:  %ROOT_DIR%
echo  Час:    %DATE% %START_TIME%
echo.
if exist "%ROOT_DIR%\setup.exe" (
    echo  setup.exe:   [OK]
) else (
    echo  setup.exe:   [немає] - пункт [1]
)
if exist "%ROOT_DIR%\Office\Data" (
    echo  Office\Data: [OK]
) else (
    echo  Office\Data: [немає] - пункт [2]
)
call :load_office_info
if defined EXISTING_OFFICE (
    echo  Office:      !EXISTING_OFFICE!
    echo  Версія:      v!OFFICE_VERSION!  ^(!OFFICE_PLATFORM!^)  !OFFICE_CHANNEL!
    if defined OFFICE_EDITION echo  Видання:     !OFFICE_EDITION!
) else (
    echo  Office:      [не встановлено]
)
echo.
exit /b

:menu
call :header
echo  Оберіть дію:
echo.
echo    --- Підготовка ^(потрібен інтернет^) ---
echo    [1] Завантажити ODT ^(setup.exe, ~2 MB^)
echo    [2] Завантажити офлайн-пакет Office ^(~2 GB^)
echo.
echo    --- Встановлення ---
echo    [3] Встановити Office 2021 Pro Plus ^(офлайн^)
echo    [4] Перевстановити Office ^(видалити + встановити^)
echo.
echo    --- Обслуговування ---
echo    [5] Перевірити стан Office
echo    [6] Повністю видалити Office з системи
echo    [7] Перевірити / видалити Skype та Teams
echo.
echo    [0] Вихід
echo.
choice /C 01234567 /N /M "  Ваш вибір (0-7): "
if errorlevel 8 goto :action_remove_teams_skype
if errorlevel 7 goto :action_uninstall
if errorlevel 6 goto :action_status
if errorlevel 5 goto :action_reinstall
if errorlevel 4 goto :action_install
if errorlevel 3 goto :action_download
if errorlevel 2 goto :action_download_odt
if errorlevel 1 goto :exit_ok
goto :menu

:ask_menu_or_exit
echo.
choice /C 01 /N /M "  [1] Головне меню  [0] Вихід: "
if errorlevel 2 goto :menu
exit /b 1

:prompt_back
echo.
echo    [1] Продовжити
echo    [0] Назад до головного меню
echo.
choice /C 01 /N /M "  Ваш вибір: "
if errorlevel 2 exit /b 0
exit /b 1

:load_office_info
set "EXISTING_OFFICE="
set "OFFICE_VERSION="
set "OFFICE_PLATFORM="
set "OFFICE_CHANNEL="
set "OFFICE_EDITION="
for /f "usebackq tokens=1,2,3,4,5 delims=|" %%A in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Office-Installer.ps1" -Command GetInstalledOffice -OutputFormat Brief 2^>nul`) do (
    set "EXISTING_OFFICE=%%~A"
    set "OFFICE_VERSION=%%~B"
    set "OFFICE_PLATFORM=%%~C"
    set "OFFICE_CHANNEL=%%~D"
    set "OFFICE_EDITION=%%~E"
)
exit /b 0

:check_setup
if not exist "%ROOT_DIR%\setup.exe" (
    echo  [X] НЕ ЗНАЙДЕНО: setup.exe
    echo      Спочатку виконайте пункт [1] - завантажити ODT
    goto :error
)
echo  [OK] setup.exe
if not exist "%ROOT_DIR%\configurations.xml" if not exist "%ROOT_DIR%\configuration.xml" (
    echo  [X] НЕ ЗНАЙДЕНО: configurations.xml
    goto :error
)
echo  [OK] configurations.xml
if not exist "%ROOT_DIR%\Office-Installer.ps1" (
    echo  [X] НЕ ЗНАЙДЕНО: Office-Installer.ps1
    goto :error
)
echo  [OK] Office-Installer.ps1
exit /b 0

:check_offline
if not exist "%ROOT_DIR%\Office\Data" (
    echo  [X] НЕ ЗНАЙДЕНО: папка Office\Data
    echo.
    echo  Спочатку виконайте пункт [2] або: Install-Office2021.bat /download
    goto :error
)
for /f "usebackq delims=" %%S in (`powershell -NoProfile -Command "(Get-ChildItem '%ROOT_DIR%\Office' -Recurse -File -EA SilentlyContinue | Measure-Object Length -Sum).Sum"`) do set "PKG_SIZE=%%S"
if not defined PKG_SIZE set "PKG_SIZE=0"
if %PKG_SIZE% LSS 100000000 (
    echo  [X] Офлайн-пакет занадто малий ^(менше 100 MB^)
    goto :error
)
echo  [OK] Office\Data ^(офлайн-файли на місці^)
exit /b 0

:check_setup_running
tasklist /FI "IMAGENAME eq setup.exe" 2>nul | find /I "setup.exe" >nul
if %errorLevel% equ 0 (
    echo  [X] Інший setup.exe уже працює. Дочекайтеся завершення.
    goto :error
)
exit /b 0

:action_download_odt
call :header
echo  [1] Завантаження Office Deployment Tool
call :prompt_back
if errorlevel 1 goto :menu
echo.
if exist "%ROOT_DIR%\setup.exe" (
    echo  [i] setup.exe уже є у цій папці.
    choice /C YN0 /N /M "  Y=завантажити знову  N=меню  0=меню: "
    if errorlevel 3 goto :menu
    if errorlevel 2 goto :menu
)
echo.
echo  Завантаження з Microsoft та розпакування setup.exe...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Office-Installer.ps1" -Command DownloadOdt -InstallDir "%ROOT_DIR%"
set "VERIFY_CODE=%ERRORLEVEL%"
echo.
if %VERIFY_CODE% equ 0 (
    echo  ============================================================
    echo    ГОТОВО - setup.exe готовий. Далі: пункт [2]
    echo  ============================================================
) else (
    echo  ============================================================
    echo    ПОМИЛКА - не вдалося отримати setup.exe
    echo  ============================================================
)
call :ask_menu_or_exit
goto :menu

:action_download
call :header
echo  [2] Завантаження офлайн-пакету Office
call :prompt_back
if errorlevel 1 goto :menu
echo.
echo  [Крок 1/2] Перевірка файлів...
echo.
call :check_setup
if errorlevel 1 goto :error
if exist "%ROOT_DIR%\Office\Data" (
    echo.
    echo  [i] Офлайн-пакет уже існує у Office\Data
    choice /C YN0 /N /M "  Y=завантажити  N=меню  0=меню: "
    if errorlevel 3 goto :menu
    if errorlevel 2 goto :menu
)
echo.
echo  [Крок 2/2] Завантаження з Microsoft CDN...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Office-Installer.ps1" -Command Download -InstallDir "%ROOT_DIR%" -ScriptDir "%SCRIPT_DIR%"
set "VERIFY_CODE=%ERRORLEVEL%"
echo.
if %VERIFY_CODE% equ 0 (
    echo  ============================================================
    echo    ГОТОВО - офлайн-пакет готовий. Далі: пункт [3]
    echo  ============================================================
    echo.
    echo  Скопіюйте всю папку:
    echo    %ROOT_DIR%
) else (
    echo  ============================================================
    echo    ПОМИЛКА - завантаження не вдалося
    echo  ============================================================
)
call :ask_menu_or_exit
goto :menu

:action_install
call :header
echo  [3] Встановлення Office 2021 Pro Plus
call :prompt_back
if errorlevel 1 goto :menu
echo.
echo  [Крок 1/4] Перевірка файлів...
echo.
call :check_setup
if errorlevel 1 goto :error
call :check_offline
if errorlevel 1 goto :error
echo.
echo  [Крок 2/4] Перевірка наявного Office...
echo.
call :check_existing_office
if errorlevel 1 goto :menu
if "%SKIP_INSTALL%"=="1" goto :verify
set "REINSTALL=0"
goto :do_install

:action_reinstall
call :header
echo  [4] Перевстановлення Office
call :prompt_back
if errorlevel 1 goto :menu
echo.
echo  [Крок 1/4] Перевірка файлів...
echo.
call :check_setup
if errorlevel 1 goto :error
call :check_offline
if errorlevel 1 goto :error
echo.
echo  [УВАГА] Буде видалено поточний Office і встановлено заново.
echo.
choice /C YN0 /N /M "  Y=так  N=ні  0=меню: "
if errorlevel 3 goto :menu
if errorlevel 2 goto :menu
set "REINSTALL=1"
goto :do_install

:action_status
call :header
echo  [5] Перевірка стану Office
echo.
if exist "%ROOT_DIR%\configurations.xml" (
    echo  [OK] configurations.xml
) else (
    echo  [X] configurations.xml
)
if exist "%ROOT_DIR%\setup.exe" (
    echo  [OK] setup.exe
) else (
    echo  [X] setup.exe - пункт [1]
)
call :load_office_info
if defined EXISTING_OFFICE (
    echo  [OK] Office встановлений:
    echo       !EXISTING_OFFICE!
    echo       Версія: v!OFFICE_VERSION!  !OFFICE_PLATFORM!  !OFFICE_CHANNEL!
    if defined OFFICE_EDITION echo       Видання: !OFFICE_EDITION!
) else (
    echo  [i] Office у системі не знайдено
)
if exist "%ROOT_DIR%\Office\Data" (
    for /f "usebackq delims=" %%S in (`powershell -NoProfile -Command "$s=(Get-ChildItem '%ROOT_DIR%\Office' -Recurse -File -EA SilentlyContinue|Measure-Object Length -Sum).Sum; if($s -ge 1GB){'{0:N2} GB'-f($s/1GB)}elseif($s -ge 1MB){'{0:N0} MB'-f($s/1MB)}else{'-'}"`) do echo  [OK] Офлайн-пакет: %%S у Office\Data
) else (
    echo  [i] Офлайн-пакет не завантажено - пункт [2]
)
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Office-Installer.ps1" -Command VerifyInstall -SetupExitCode 0 -AlreadyInstalled 2>nul
echo.
call :ask_menu_or_exit
goto :menu

:action_uninstall
call :header
echo  [6] Повне видалення Office
call :prompt_back
if errorlevel 1 goto :menu
echo.
echo  [Крок 1/3] Перевірка файлів...
echo.
call :check_setup
if errorlevel 1 goto :error
echo.
call :check_existing_office_silent
if not defined EXISTING_OFFICE (
    echo  [i] Office у системі не знайдено - видаляти нічого.
    call :ask_menu_or_exit
    goto :menu
)
echo  [УВАГА] Буде видалено ВСІ продукти Microsoft Office:
echo          !EXISTING_OFFICE!
echo          Версія: v!OFFICE_VERSION!  !OFFICE_PLATFORM!  !OFFICE_CHANNEL!
if defined OFFICE_EDITION echo          Видання: !OFFICE_EDITION!
echo.
choice /C YN0 /N /M "  Y=так  N=ні  0=меню: "
if errorlevel 3 goto :menu
if errorlevel 2 goto :menu
echo.
echo  [Крок 2/3] Видалення Office...
echo    Тихий режим — прогрес у цьому вікні.
echo    Початок: %TIME%
echo    --------------------------------------------------------
call :check_setup_running
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Office-Installer.ps1" -Command Uninstall -InstallDir "%ROOT_DIR%"
set "EXIT_CODE=%ERRORLEVEL%"
echo    --------------------------------------------------------
echo    Завершено: %TIME%  ^(код: %EXIT_CODE%^)
echo.
echo  [Крок 3/3] Перевірка результату...
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Office-Installer.ps1" -Command VerifyRemoved -SetupExitCode %EXIT_CODE%
set "VERIFY_CODE=%ERRORLEVEL%"
echo.
if %VERIFY_CODE% equ 0 (
    echo  ============================================================
    echo    ГОТОВО - Office повністю видалено
    echo  ============================================================
) else (
    echo  ============================================================
    echo    УВАГА - Office може бути видалено не повністю
    echo  ============================================================
)
call :ask_menu_or_exit
goto :menu

:action_remove_teams_skype
call :header
echo  [7] Skype та Microsoft Teams
call :prompt_back
if errorlevel 1 goto :menu
echo.
echo  [Крок 1/2] Перевірка...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Office-Installer.ps1" -Command CheckTeamsSkype
set "TS_CODE=%ERRORLEVEL%"
echo.
if %TS_CODE% equ 0 (
    call :ask_menu_or_exit
    goto :menu
)
echo  [УВАГА] Буде видалено знайдені Skype / Teams.
echo         Office ^(Word, Excel тощо^) залишиться без змін.
echo.
choice /C YN0 /N /M "  Y=видалити  N=меню  0=меню: "
if errorlevel 3 goto :menu
if errorlevel 2 goto :menu
echo.
echo  [Крок 2/2] Видалення ^(може тривати кілька хвилин^)...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Office-Installer.ps1" -Command RemoveTeamsSkype -InstallDir "%ROOT_DIR%"
set "TS_CODE=%ERRORLEVEL%"
echo.
if %TS_CODE% equ 0 (
    echo  ============================================================
    echo    ГОТОВО - Skype та Teams видалено
    echo  ============================================================
) else (
    echo  ============================================================
    echo    УВАГА - видалено не повністю, див. повідомлення вище
    echo  ============================================================
)
call :ask_menu_or_exit
goto :menu

:check_existing_office
call :load_office_info
if defined EXISTING_OFFICE (
    echo  [УВАГА] Office уже встановлений:
    echo          !EXISTING_OFFICE!
    echo          Версія: v!OFFICE_VERSION!  !OFFICE_PLATFORM!  !OFFICE_CHANNEL!
    if defined OFFICE_EDITION echo          Видання: !OFFICE_EDITION!
    echo.
    echo    [1] Вийти без змін ^(до меню^)
    echo    [2] Перевстановити ^(пункт [4]^)
    echo    [0] Назад до головного меню
    echo.
    choice /C 012 /N /M "  Ваш вибір (0-2): "
    if errorlevel 3 (
        set "REINSTALL=1"
        goto :do_install
    )
    if errorlevel 2 (
        set "SKIP_INSTALL=1"
        exit /b 0
    )
    exit /b 1
)
exit /b 0

:check_existing_office_silent
call :load_office_info
exit /b 0

:do_install
echo.
if "%REINSTALL%"=="1" (
    echo  [Крок 3-4/4] Перевстановлення Office...
    echo    A - видалення поточної версії
    echo    B - встановлення з офлайн-файлів
    echo    Прогрес відображається у цьому вікні.
) else (
    echo  [Крок 3-4/4] Встановлення Office...
    echo    Тихий режим Microsoft - прогрес у цьому вікні.
)
echo    Початок: %TIME%
echo    --------------------------------------------------------
call :check_setup_running
if "%REINSTALL%"=="1" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Office-Installer.ps1" -Command Install -InstallDir "%ROOT_DIR%" -Mode Reinstall
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Office-Installer.ps1" -Command Install -InstallDir "%ROOT_DIR%" -Mode Install
)
set "EXIT_CODE=%ERRORLEVEL%"
echo    --------------------------------------------------------
echo    Завершено: %TIME%  ^(код: %EXIT_CODE%^)
echo.
goto :verify

:verify
echo  [Підсумок] Перевірка результату...
echo.
if "%SKIP_INSTALL%"=="1" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Office-Installer.ps1" -Command VerifyInstall -SetupExitCode 0 -AlreadyInstalled
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Office-Installer.ps1" -Command VerifyInstall -SetupExitCode %EXIT_CODE%
)
set "VERIFY_CODE=%ERRORLEVEL%"
echo.
if "%SKIP_INSTALL%"=="1" (
    echo  ============================================================
    echo    Office УЖЕ ВСТАНОВЛЕНИЙ - без змін
    echo  ============================================================
) else if %VERIFY_CODE% equ 0 (
    if "%REINSTALL%"=="1" (
        echo  ============================================================
        echo    ГОТОВО - Office перевстановлено
        echo  ============================================================
    ) else (
        echo  ============================================================
        echo    ГОТОВО - Office встановлено
        echo  ============================================================
    )
    echo.
    echo  Активація: Word/Excel -^> Файл -^> Обліковий запис -^> Активувати Office
) else if %VERIFY_CODE% equ 3 (
    echo  ============================================================
    echo    Office уже є - спробуйте пункт [4]
    echo  ============================================================
) else (
    echo  ============================================================
    echo    ПОМИЛКА - встановлення не вдалося ^(код %EXIT_CODE%^)
    echo  ============================================================
)
call :ask_menu_or_exit
goto :menu

:error
echo.
echo  ============================================================
echo    ПОМИЛКА
echo  ============================================================
call :ask_menu_or_exit
goto :menu

:exit_ok
echo.
echo  До побачення.
exit /b 0
