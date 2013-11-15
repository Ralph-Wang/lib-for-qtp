Option Explicit

Private Const LogName = "test.log" ' where to write

'test
Private Sub test()
    writeLog("write a log")
End Sub

Public Function writeLog(logMsg):
    Dim localTime, NewLog, logFile
    localTime = "[" & Now() & "]"
    NewLog = localTime & " " & logMsg
    Set logFile = openLogFile()
    logFile.WriteLine(NewLog)
    Set logFile = Nothing
End Function

Private Function openLogFile()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(LogName) Then
        Set openLogFile = fso.OpenTextFile(LogName, 8, false)
    Else
        fso.CreateTextFile(LogName)
        Set openLogFile = fso.OpenTextFile(LogName, 8, false)
    End If
    Set fso = Nothing
End Function
