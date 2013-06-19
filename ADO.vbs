'QTP(UFT)用数据库连接库
Option Explicit

'测试函数,使用前请注释掉
'test
Private Function test()
    Const DB = "test"
    Dim DBstr_test
    DBstr_test = "Provider=SQLNCLI10;Server=(local);Database=" & DB & ";Uid=sa;Pwd="
    Dim res
    'Set res = SQL_Output1Line(DBstr, "select * from test")
    'msgbox SQL_GetValue1line(res, "a")

    'Set res = SQL_Output1Field(DBstr, "select a from test")
    'msgbox SQL_GetValue1Field(res, 1)

    Set res = SQL_OutputFullTable(DBstr_test, "select * from test")
    msgbox SQL_GetValueFullTable(res, 0, "a")
End Function

'****************************************************************************************************************************
'函数名:SQL_GetValue1line
'功能:与SQL_Output1Line配合使用,获取其返回结果中的具体值
'参数:Dict, SQL_Output1Line的返回对象; key,SQL_Output1Line查询结果的字段名(别名)
'****************************************************************************************************************************
Public Function SQL_GetValue1Line(Dict, key)
    SQL_GetValue1Line = Dict.Item(key)
End Function

'****************************************************************************************************************************
'函数名:SQL_GetValue1Field
'功能:与SQL_Output1Field配合使用,获取其返回结果中的具体值
'参数:Dict, SQL_Output1Field的返回对象; rowNum,SQL_Output1Field查询结果的行号(从0开始)
'****************************************************************************************************************************
Public Function SQL_GetValue1Field(Dict, rowNum)
    Dim key, i
    key = -1
    For each i in Dict.Keys()
        If InStr(1, i, Cstr(lineNo)) Then
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

'****************************************************************************************************************************
'函数名:SQL_GetValueFullTable
'功能:与SQL_OutputFullTable配合使用,获取其返回结果中的具体值
'参数:Dict, SQL_Output1Field的返回对象; rowNum,SQL_Output1Field查询结果的行号(从0开始); key,SQL_OutputFullTable查询结果中的字段名
'****************************************************************************************************************************
Public Function SQL_GetValueFullTable(Dict, rowNum, key)
    SQL_GetValueFullTable = Dict.Item(Cstr(rowNum)).Item(key)
End Function


'****************************************************************************************************************************
'函数名:SQL_Output1Field
'功能:执行SQL将结果保存到Dict对象中并返回,只包含查询结果中第一个字段的所有值
'参数:DBstr, 数据库连接串(可以是ODBC名); strSQL, 需要执行的SQL语句
'****************************************************************************************************************************
Public Function SQL_Output1Field(DBstr, strSQL)
    Dim objCn, objRe
    Set objCn = GetDBConnection(DBstr)
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

'****************************************************************************************************************************
'函数名:SQL_Output1Line
'功能:执行SQL将结果保存到Dict对象中并返回,只包含查询结果中所有字段的第一行值
'参数:DBstr, 数据库连接串(可以是ODBC名); strSQL, 需要执行的SQL语句
'****************************************************************************************************************************
Public Function SQL_Output1Line(DBstr, strSQL)
    Dim objCn, objRe
    Set objCn = GetDBConnection(DBstr)
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

'****************************************************************************************************************************
'函数名:SQL_OutputFullTable
'功能:执行SQL将结果保存到Dict对象中并返回,包含查询结果中的全部结果
'参数:DBstr, 数据库连接串(可以是ODBC名); strSQL, 需要执行的SQL语句
'****************************************************************************************************************************
Public Function SQL_OutputFullTable(Dbstr, strSQL)
    Dim objCn, objRe
    Set objCn = GetDBConnection(DBstr)
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

'****************************************************************************************************************************
'函数名:execSQL
'功能:执行strSQL并返回一个RecordSet对象
'参数:objCn, ADODB连接对象; strSQL, 需要执行的SQL语句
'****************************************************************************************************************************
Private Function execSQL(objCn, strSQL)
    Dim objRe, save_strSQL
    save_strSQL = strSQL
    Set objRe = CreateObject("ADODB.RecordSet")
    objRe.Open save_strSQL, objcn, 0
    Set execSQL = objRe
    Set objRe = Nothing
End Function

'****************************************************************************************************************************
'函数名:GetDBConnection
'功能:返回一个连接DBstr的ADO连接对象
'参数:DBstr, 数据库连接串(ODBC名)
'****************************************************************************************************************************
Private Function GetDBConnection(DBstr)
    Dim objCn
    Set objCn = CreateObject("ADODB.Connection")
    objCn.Open DBstr
    Set GetDBConnection = objCn
    Set objCn = Nothing
End Function
