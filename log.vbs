Option Explicit

Private Const LogName = "test.log" ' where to write

'test
Private Sub test()
    WriteLog("write a log")
End Sub

Public Function WriteLog(logMsg):
    Dim localTime, NewLog, logFile
    localTime = "[" & Now() & "]"
    NewLog = localTime & " " & logMsg
    Set logFile = OpenLogFile()
    logFile.WriteLine(NewLog)
    Set logFile = Nothing
End Function

Private Function OpenLogFile()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(LogName) Then
        Set OpenLogFile = fso.OpenTextFile(LogName, 8, false)
    Else
        fso.CreateTextFile(LogName)
        Set OpenLogFile = fso.OpenTextFile(LogName, 8, false)
    End If
    Set fso = Nothing
End Function
