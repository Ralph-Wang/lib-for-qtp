Option Explicit
test
Function test()
    Dim t
    Set t = GetConfigFromExcel("test.xls")
    Dim k
    k = t.keys
    Dim i
    For i = 0 to t.Count - 1
        msgbox k(i) & vbcrlf & t.item(k(i))
    Next
End Function
Function GetConfigFromExcel(FileName)
    Dim ExcelString
    ExcelString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileName & ";Extended Properties=Excel 12.0"
    Dim objCn, objExcel, strSQL
    Set objCn = CreateObject("ADODB.Connection")
    Set objExcel = CreateObject("ADODB.RecordSet")
    objCn.Open ExcelString
    strSQL = "SELECT Item, Value from [Config$]"
    objExcel.Open strSQL, objCn, 0
    Dim Dict
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = 0
    Do Until objExcel.EOF
        Dict.Add cstr(objExcel("Item")), cstr(objExcel("Value"))
        objExcel.movenext
    Loop
    Set GetConfigFromExcel = Dict
End Function
