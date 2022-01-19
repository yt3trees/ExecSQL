@echo off

echo 右クリックコンテキストメニューの「送る」で使用できるようにします。
echo ■ExecSQL.exe

if exist %APPDATA%\Microsoft\Windows\SendTo\ExecSQL.exe (
rem    echo 既に存在するため上書きします。
rem    echo A | xcopy .\ExecSQL.exe %APPDATA%\Microsoft\Windows\SendTo\
    xcopy .\ExecSQL.exe %APPDATA%\Microsoft\Windows\SendTo\
) else (
    echo 新規でコピーします。
rem    echo F | xcopy .\ExecSQL.exe %APPDATA%\Microsoft\Windows\SendTo\
    xcopy .\ExecSQL.exe %APPDATA%\Microsoft\Windows\SendTo\
)

echo ■ExecSQL.config.json
if exist %APPDATA%\Microsoft\Windows\SendTo\ExecSQL.config.json (
rem    echo 既に存在するため上書きします。
rem    echo A | xcopy .\ExecSQL.config.json %APPDATA%\Microsoft\Windows\SendTo\
    xcopy .\ExecSQL.config.json %APPDATA%\Microsoft\Windows\SendTo\
) else (
    echo 新規でコピーします。
rem    echo F | xcopy .\ExecSQL.config.json %APPDATA%\Microsoft\Windows\SendTo\
    xcopy .\ExecSQL.config.json %APPDATA%\Microsoft\Windows\SendTo\
)

echo インストール完了しました。

pause