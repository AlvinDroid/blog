Function DoTest()
  On Error Resume Next 
	Dim a,b,c
	a=1:b=0
	c=a/b
	Set ObjLog=New SystemLog
	If Err.Number<>0 Then 
		ObjLog.WriteLog (Err.Description&"   "&Err.Number&"    "&Err.Source)
	End If
End Function 



Class SystemLog
	Private fos
	Private Sub Class_Initialize
    	Set fos = CreateObject("scripting.filesystemObject")
    End Sub

    '
    Private Sub Class_Terminate
        Set fos = Nothing
    End Sub
    
    Public Function  WriteLog(strLog)
    	Const ForReading=1,ForWriting=2,ForAppending=8
    	currentPath=createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path    	MsgBox "3"
    	currentPath=currentPath&"\SystemLog.log"
    	strLog=FormatDateTime(now,0)&"     "&strLog
    	If (fos.FileExists(currentPath)) Then
    		Set f=fos.OpenTextFile(currentPath,ForAppending)
    		f.WriteLine(strLog)
    		Set f=Nothing 
    	Else 
    		Set f=fos.CreateTextFile(currentPath)
    		f.WriteLine(strLog)
    		Set f=Nothing 
    	End If 

    End Function
End Class


DoTest()
