Option Explicit

'test ' 测试函数,使用前请注释
Function test()
    Dim t
    Set t = GetConfigFromExcel("url.xlsx","PM", "page", "url")
    Dim k
    k = t.keys
    Dim i
    For i = 0 to t.Count - 1
        msgbox k(i) & vbcrlf & t.item(k(i))
    Next
End Function


'*******************************************************
'函数名:GetConfigFromExcel
'功能:从xls文件中取值
'参数:FileName,文件名; Sheet,表单名; 
'     sItem, 第一列 (一般作为配置项名称);
'     sValue, 第二列(一般作为配置项值)
'*******************************************************
Function GetConfigFromExcel(FileName, Sheet, sItem, sValue)
    Dim ExcelString
    ExcelString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileName & ";Extended Properties=Excel 12.0"
    'ExcelString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileName & ";Extended Properties=Excel 8.0"
    Dim objCn, objExcel, strSQL
    Set objCn = CreateObject("ADODB.Connection")
    Set objExcel = CreateObject("ADODB.RecordSet")
    objCn.Open ExcelString
    strSQL = "SELECT " & sItem & ", " & sValue & " from [" & Sheet &"$]"
    objExcel.Open strSQL, objCn, 0
    Dim Dict
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = 0
    Do Until objExcel.EOF
		'msgbox objExcel(sItem)
        Dict.Add cstr(objExcel(sItem)), cstr(objExcel(sValue))
        objExcel.movenext
    Loop
    Set GetConfigFromExcel = Dict
End Function
