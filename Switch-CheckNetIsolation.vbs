Select Case msgbox("��� '��' ʹ����WindowsApp���Դ���������磬" & vbcrlf & vbcrlf & "��� '��' �������WindowsApp�����������Ȩ�ޣ�", 51, "ʹWindowsApp��ȡ����Ȩ��")  
Case 6  
    Set ws = WScript.CreateObject("wscript.shell")  
    app = ws.ExpandEnvironmentStrings("%USERPROFILE%\AppData\Local\Packages")  
    Set fso = WScript.CreateObject("scripting.filesystemobject")  
    Set fs = fso.GetFolder(app).SubFolders  
    Set bat = fso.createtextfile(ws.ExpandEnvironmentStrings("%TEMP%\WindowsAppProxyAccess.bat"))  
    For Each f In fs  
        bat.WriteLine ("CheckNetIsolation.exe LoopbackExempt -a -n=" & f.name)  
    Next  
    bat.WriteLine ("del %0")  
    CreateObject("Shell.Application").ShellExecute ws.ExpandEnvironmentStrings("%TEMP%\WindowsAppProxyAccess.bat"),"","","runas",1  
Case 7  
    CreateObject("Shell.Application").ShellExecute "CheckNetIsolation.exe","LoopbackExempt -c","","runas",1  
End Select  