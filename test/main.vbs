include "log.vbs"
'writeLog "------------------------------"
'writeLog "test Start"
'''''''''''
' > test suits
'''''''''''
'writeLog "test Ended"
'------------------------------------------------------------
' include other *.vbs files
Function include(strFile)
  Dim fso, file, content
  Set fso = createObject("Scripting.FileSystemObject")
  strFile = fso.getAbsolutePathName(strFile)
  Set file = fso.openTextFile(strFile)
  msgbox strFile
  content = file.readAll()
  file.close
  msgbox content
  ExecuteGlobal content
  set fso = Nothing
  set file = Nothing
End Function

' an implementation of assert
Function assert(expression, errDescription)
  If not expression Then
    err.raise 10000,"AssertError", errDescription
  End If
End Function
