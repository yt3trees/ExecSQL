rem upx --upx-dir C:\upx-3.96-win64
rem --icon=icon/copytool.ico
pyinstaller ExecSQL.py --name ExecSQL --noconsole --icon=icon/ExecSQL_LightIcon.ico

if exist .\app\ExecSQL\ExecSQL.exe (
    echo �㏑�����܂��B
    echo A | xcopy .\dist .\app /S
) else (
    echo �V�K�ŃR�s�[���܂��B
    echo D | xcopy .\dist .\app /S
)
pause