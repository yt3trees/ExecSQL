@echo off

echo �E�N���b�N�R���e�L�X�g���j���[�́u����v�Ŏg�p�ł���悤�ɂ��܂��B
echo ��ExecSQL.exe

if exist %APPDATA%\Microsoft\Windows\SendTo\ExecSQL.exe (
rem    echo ���ɑ��݂��邽�ߏ㏑�����܂��B
rem    echo A | xcopy .\ExecSQL.exe %APPDATA%\Microsoft\Windows\SendTo\
    xcopy .\ExecSQL.exe %APPDATA%\Microsoft\Windows\SendTo\
) else (
    echo �V�K�ŃR�s�[���܂��B
rem    echo F | xcopy .\ExecSQL.exe %APPDATA%\Microsoft\Windows\SendTo\
    xcopy .\ExecSQL.exe %APPDATA%\Microsoft\Windows\SendTo\
)

echo ��ExecSQL.config.json
if exist %APPDATA%\Microsoft\Windows\SendTo\ExecSQL.config.json (
rem    echo ���ɑ��݂��邽�ߏ㏑�����܂��B
rem    echo A | xcopy .\ExecSQL.config.json %APPDATA%\Microsoft\Windows\SendTo\
    xcopy .\ExecSQL.config.json %APPDATA%\Microsoft\Windows\SendTo\
) else (
    echo �V�K�ŃR�s�[���܂��B
rem    echo F | xcopy .\ExecSQL.config.json %APPDATA%\Microsoft\Windows\SendTo\
    xcopy .\ExecSQL.config.json %APPDATA%\Microsoft\Windows\SendTo\
)

echo �C���X�g�[���������܂����B

pause