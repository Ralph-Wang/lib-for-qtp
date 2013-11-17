Function include(strFile)
  Dim fso, file, content
  Set fso = createObject("Scripting.FileSystemObject")
  Set file = fso.openTextFile(strFile)
  content = file.readAll()
  file.close
  ExecuteGlobal content
End Function

Function assert(expression, errDescription)
  If not expression Then
    err.raise 10000,"AssertError", errDescription
  End If
End Function
