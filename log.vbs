Option Explicit

Private Const logFileName = "test.log" ' where to write
Private Const forAppending = 8 ' open files for append

Public Function writeLog(logMsg):
    Dim localTime, logFile
    localTime = "[" & Now() & "]:"
    logMsg = localTime & " " & logMsg
    Set logFile = openLogFile()
    logFile.WriteLine(logMsg)
    Set logFile = Nothing
End Function

Private Function openLogFile()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(logFileName) Then
        Set openLogFile = fso.OpenTextFile(logFileName, forAppending, false)
    Else
        fso.CreateTextFile(logFileName)
        Set openLogFile = fso.OpenTextFile(logFileName, forAppending, false)
    End If
    Set fso = Nothing
End Function
