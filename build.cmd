rem upx --upx-dir C:\upx-3.96-win64
rem --icon=icon/copytool.ico
pyinstaller ExecSQL.py --name ExecSQL --noconsole

if exist .\app\ExecSQL\ExecSQL.exe (
    echo 上書きします。
    echo A | xcopy .\dist .\app /S
) else (
    echo 新規でコピーします。
    echo D | xcopy .\dist .\app /S
)
pause