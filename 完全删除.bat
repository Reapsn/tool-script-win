@ECHO OFF
CLS
color 0a
echo 确定要尝试删除吗?按任意键开始删除！
pause >nul
taskkill /f /im explorer.exe>nul
echo y|Cacls %* /t /e /c /g Everyone:f >nul
DEL /F /A /Q \\?\%1 
RD /S /Q \\?\%1 
start %windir%\explorer.exe
exit