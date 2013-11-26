Const forAppending = 8 ' open files for append
Const forReading = 1 ' open files for Read
Dim g_CurrentPath, g_ScriptsPath, g_ReportsPath, g_CurrentRunTime
Dim g_ConfigFile, g_LogFileName

' init variables
g_CurrentRunTime = getCurrentRunTime()
g_CurrentPath = getCurrentPath()
g_ScriptsPath = g_CurrentPath & "\scripts"
g_ReportsPath = g_CurrentPath & "\reports\" & g_CurrentRunTime

g_ConfigFile = ".\config\running.ini" ' where to get configs
g_LogFileName = ".\log\" & g_CurrentRunTime & ".log" ' where to write log


call Main()

'''''''''''''''''''''''''''''''''''''''
' >> Main
'''''''''''''''''''''''''''''''''''''''
Function Main()
  writeLog "Create folder for reports:" & g_CurrentRunTime
  createFolderSafe(g_ReportsPath)
  writeLog "start Running"
  call runTestSuites()
  writeLog "All Suites Running Done"
  msgbox "[ " & now() & " ]:All Suites Running Done"
End Function

Function runTestSuites()
  Dim app, qtTest, qtResult
  Dim curTest, theConfigDict
  ' get config
  Set theConfigDict = getAllConfig()
  Set app = createObject("QuickTest.Application")
  'Active the 'Web' Addin
  If app.setActiveAddins(Array("Web")) Then
    writeLog "Succ to Activate Web Addins"
  Else
    writeLog "Fail to Activate Web Addins, Quit"
    set app = Nothing
    Exit Function
  End If
  'Do not open the GUI
  app.Visible = False
  app.launch
  'Open the test in Read-only Mode
  For Each curTest in split(theConfigDict("tests"), ",")
    writeLog "open test: " & curTest
    app.open g_ScriptsPath & "\" & curTest, True 
    app.Options.Run.RunMode = "Fast"
    Set qtTest = app.Test
    Set qtResult = createObject("QuickTest.RunResultsOptions")
    qtResult.ResultsLocation = g_ReportsPath & "\" & curTest
    writeLog "start to run test: " & curTest
    qtTest.run qtResult
    writeLog "done to run test: " & curTest
  Next
  app.quit
  Set app = Nothing
  Set qtTest = Nothing
  set qtResult = Nothing
End Function

'''''''''''''''''''''''''''''''''''''''
' >> config
'''''''''''''''''''''''''''''''''''''''

' get all config items
Function getAllConfig()
  Dim theConfigFile, configSource, keyMap, configDict
  Set configDict = createObject("Scripting.Dictionary")
  Set theConfigFile = openConfigFile()
  Do While theConfigFile.atEndOfLine <> True
    configSource = theConfigFile.readLine()
    keyMap = split(trim(configSource), "=")
    configDict.add keyMap(0), keyMap(1)
  Loop
  Set getAllConfig = configDict
  Set theConfigFile = Nothing
  Set configDict = Nothing
End Function

' open g_ConfigFile
Function openConfigFile()
  Dim fso
  Set fso = createObject("Scripting.FileSystemObject")
  If fso.fileExists(g_ConfigFile) Then
    Set openConfigFile = fso.openTextFile(g_ConfigFile, forReading, false)
  Else
    Err.raise 10000, "未找到配置文件:" & g_ConfigFile, "未找到配置文件"
  End If
  Set fso = nothing
End Function


'''''''''''''''''''''''''''''''''''''''
' >> logging
'''''''''''''''''''''''''''''''''''''''
' write logMsg to logFile
Function writeLog(logMsg):
  Dim localTime, logFile
  localTime = "[" & now() & "]:"
  logMsg = localTime & " " & logMsg
  Set logFile = openLogFile()
  logFile.WriteLine(logMsg)
  Set logFile = Nothing
End Function

' open logFile for append
Function openLogFile()
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(g_LogFileName) Then
    Set openLogFile = fso.OpenTextFile(g_LogFileName, forAppending, false)
  Else
    fso.CreateTextFile(g_LogFileName)
    Set openLogFile = fso.OpenTextFile(g_LogFileName, forAppending, false)
  End If
  Set fso = Nothing
End Function


'''''''''''''''''''''''''''''''''''''''
' >> others
'''''''''''''''''''''''''''''''''''''''
Function getCurrentPath()
  getCurrentPath = createObject("Scripting.FileSystemObject").getFolder(".").Path
End Function

Function getCurrentRunTime()
  Dim theRegExp
  Set theRegExp = New RegExp
  theRegExp.global = True 'Search all the target string
  theRegExp.ignoreCase = False
  theRegExp.pattern = "\D+"
  getCurrentRunTime = theRegExp.replace(Now(),"")
End Function

Function createFolderSafe(path)
  Dim fso
  Set fso = createObject("Scripting.FileSystemObject")
  If not fso.folderExists(path) Then
    fso.createFolder(path)
  End If
End Function
