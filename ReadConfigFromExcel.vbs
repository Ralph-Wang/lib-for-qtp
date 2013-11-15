Option Explicit

'test ' 测试函数,使用前请注释
Function test()
    Dim t
    Set t = getConfigFromExcel("url.xlsx","PM", "page", "url")
    Dim k
    k = t.keys
    Dim i
    For i = 0 to t.Count - 1
        msgbox k(i) & vbcrlf & t.item(k(i))
    Next
End Function


'*******************************************************
'函数名:getConfigFromExcel
'功能:从xls文件中取值
'参数:FileName,文件名; Sheet,表单名; 
'     sItem, 第一列 (一般作为配置项名称);
'     sValue, 第二列(一般作为配置项值)
'*******************************************************
Function getConfigFromExcel(FileName, Sheet, sItem, sValue)
    Dim excelString
    excelString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileName & ";Extended Properties=Excel 12.0"
    'excelString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileName & ";Extended Properties=Excel 8.0"
    Dim objCn, objExcel, strSQL
    Set objCn = CreateObject("ADODB.Connection")
    Set objExcel = CreateObject("ADODB.RecordSet")
    objCn.Open excelString
    strSQL = "SELECT " & sItem & ", " & sValue & " from [" & Sheet &"$]"
    objExcel.Open strSQL, objCn, 0
    Dim dict
    Set dict = CreateObject("Scripting.dictionary")
    dict.CompareMode = 0
    Do Until objExcel.EOF
		'msgbox objExcel(sItem)
        dict.Add cstr(objExcel(sItem)), cstr(objExcel(sValue))
        objExcel.movenext
    Loop
    Set getConfigFromExcel = dict
End Function
