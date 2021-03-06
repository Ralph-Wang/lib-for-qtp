Option Explicit

' get the very value from SQL_GetValue1Field
Public Function SQL_GetValue1Field(Dict, rowNum)
    Dim key, i
    key = -1
    For each i in Dict.Keys()
        If InStr(1, i, Cstr(rowNum)) Then
            key = i
            Exit For
        End If
    Next
    If Key = -1 Then
        MsgBox "SQL_GetValue1Field参数溢出,rowNum大于Dict最大值"
        SQL_GetValue1Field = ""
        Exit Function
    End If
    SQL_GetValue1Field = Dict.Item(key)
End Function

' get the very value from SQL_Output1Line
Public Function SQL_GetValue1Line(Dict, key)
    SQL_GetValue1Line = Dict.Item(key)
End Function



' get the very value from SQL_OutputFullTable
Public Function SQL_GetValueFullTable(Dict, rowNum, key)
    SQL_GetValueFullTable = Dict.Item(Cstr(rowNum)).Item(key)
End Function


' execute the strSQL in strDB, return the first fields in Dictionary
' hint, use SQL_GetValue1Field to get the values 
Public Function SQL_Output1Field(strDB, strSQL)
    Dim objCn, objRe
    Set objCn = getDBConnection(strDB)
    Set objRe = execSQL(objCn, strSQL)
    Dim Dict, FirstField, iter
    Set FirstField = objRe.Fields(0)
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = 0
    iter = 0
    Do While not objRe.EOF
        Dict.Add Cstr(FirstField.Name) & "_" & Cstr(iter), Cstr(FirstField.Value)
        objRe.MoveNext
        iter = iter + 1
    Loop
    Set SQL_Output1Field = Dict
    Set FirstField = Nothing
    Set objCn = Nothing
    Set objRe = Nothing
End Function

' execute the strSQL in strDB, return the first line in Dictionary
' hint, use SQL_GetValue1Line to get the values 
Public Function SQL_Output1Line(strDB, strSQL)
    Dim objCn, objRe
    Set objCn = getDBConnection(strDB)
    Set objRe = execSQL(objCn, strSQL)
    Dim Dict, iter
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = 0
    For each iter in objRe.Fields
        Dict.Add Cstr(iter.Name), Cstr(iter.Value)
    Next
    Set SQL_Output1Line = Dict
    Set objCn = Nothing
    Set objRe = Nothing
    Set Dict = Nothing
End Function


' execute the strSQL in strDB, return the full table in Dictionary
' hint, use SQL_GetValueFullTable to get the values 
Public Function SQL_OutputFullTable(strDB, strSQL)
    Dim objCn, objRe
    Set objCn = getDBConnection(strDB)
    Set objRe = execSQL(objCn, strSQL)
    Dim DictFields, DictLines, iter, jter, Fields, iter_Name, iter_Value
    Set DictLines = CreateObject("Scripting.Dictionary")
    DictLines.CompareMode = 0
    jter = 0
    Set Fields = objRe.Fields
    Do While not objRe.EOF
        Set DictFields = CreateObject("Scripting.Dictionary")
        DictFields.CompareMode = 0
        For each iter in Fields
            iter_Name = iter.Name
            if isNull(iter.Value) then
                iter_Value = ""
            Else
                iter_Value = iter.Value
            End if
            DictFields.Add Cstr(iter_Name), Cstr(iter_Value)
        Next
        DictLines.Add Cstr(jter), DictFields
        jter = jter + 1
        objRe.MoveNext
        Set DictFields = Nothing
    Loop
    Set SQL_OutputFullTable = DictLines
    Set objCn = Nothing
    Set objRe = Nothing
    Set DictFields = Nothing
    Set DictLines = Nothing
End Function

' execute the strSQL in connection objCn
Private Function execSQL(objCn, strSQL)
    Dim objRe, save_strSQL
    save_strSQL = strSQL
    Set objRe = CreateObject("ADODB.RecordSet")
    objRe.Open save_strSQL, objcn, 0
    Set execSQL = objRe
    Set objRe = Nothing
End Function

' get the DB connection of strDB
Private Function getDBConnection(strDB)
    Dim objCn
    Set objCn = CreateObject("ADODB.Connection")
    objCn.Open strDB
    Set getDBConnection = objCn
    Set objCn = Nothing
End Function
