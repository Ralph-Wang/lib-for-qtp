Option Explicit

'test
Private Function test()
    Dim t
    Set t = GetConfigFromExcel("C:\Users\Administrator\Desktop\github\lib-for-qtp\test.xls","test", "sItem", "sValue")
    Dim k
    k = t.keys
    Dim i
    For i = 0 to t.Count - 1
        msgbox k(i) & vbcrlf & t.item(k(i))
    Next
End Function

' get sItem&sValue as Dictionary from sheetName of fileName
' unstable for x86
Public Function getConfigFromExcel(fileName, sheetName, sItem, sValue)
    Dim ExcelString
    'ExcelString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileName & ";Extended Properties=Excel 12.0"
    ExcelString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fileName & ";Extended Properties=Excel 8.0"
    Dim objCn, objExcel, strSQL
    Set objCn = CreateObject("ADODB.Connection")
    Set objExcel = CreateObject("ADODB.RecordSet")
    objCn.Open ExcelString
    strSQL = "SELECT " & sItem & ", " & sValue & " from [" & sheetName &"$]"
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
