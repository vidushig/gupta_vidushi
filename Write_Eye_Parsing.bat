@set path=%PATH%;\\teentaal\vi_tools\sve_tools_python\Python37 \Lib\site-packages;\\teentaal\vi_tools\sve_tools_python\Python37\Lib

@set path=%PATH%;C:\Windows\System32\downlevel
@set PYTHONPATH=%PATH%;%cd%\..\;%cd%\..\..\

@cd C:\Users\c_vidush\Documents\Qranium\Data_Parsing\

@pyinstaller --win-private-assemblies --win-no-prefer-redirects -w -F -n Data_Parsing.exe Data_Parsing.py

pause
::C:\Users\c_vidush\Documents\Qranium\Data_Parsing\dist\Data_Parsing\Data_Parsing.exe


