Option Explicit

'*******************************************************
'函数名:regSearch
'功能:通过正则表达式查找, 返回第一个匹配值
'参数:thePattern, 正则模式;
'     strToSearch, 需要查找的字符串
'*******************************************************
Function regSearch(thePattern, strToSearch)
	Dim RegEx,  Matches
	Set regEx = New RegExp
	regEx.Pattern = thePattern
	regEx.IgnoreCase = False
	regEx.Global = false
	Set Matches = regEx.Execute(strToSearch)
    regSearch = Matches(0).Value
End Function
'*******************************************************
'函数名:checkTest
'功能:检查测试实际值与用例期望值是否一致; 一致则返回True,否则返回False
'参数:expvalue,用例期望值 ; actvalue, 测试实际值   
'     method, 0代表精确匹配, 1代表模糊匹配
'*******************************************************
public Function checkTest(expvalue, actvalue, method)
	Dim res
	If method <> 0 and method <> 1 Then
		msgbox "checkTest函数的method参数不正确; 0-精确匹配, 1-模糊匹配"
	Else
		Select Case method
			Case 0
				res = (expvalue = actvalue)
				checkTest = res
			Case 1
				res = (Instr(1,expvalue,actvalue) + Instr(1,actvalue,expvalue) > 0)
				checkTest = res
			Case Else
		End Select
	End If
End Function

'*******************************************************
'函数名:myReporter
'功能:根据测试结果输出标准测试报告; 无返回值
'参数:tcase,用例名称 ; res,布尔值,测试结果
'*******************************************************
public Function myReporter(tcaseWithValues, res)
	If typename(res) = "Boolean" Then
		If res Then
			reporter.ReportEvent micPass, tcaseWithValues, "测试通过"
		Else
			reporter.ReportEvent micFail, tcaseWithValues, "测试未通过"
		End If
	Else
		msgbox "错误的res参数("& res &"),请输入bool类型参数"
	End If
End Function

'*******************************************************
'函数名:verifyDBres
'功能:根据测试结果输出标准测试报告; 无返回值
'参数:tcase,用例名称 ; res,布尔值,测试结果
'*******************************************************
public Function verifyDBvalue(expvalue, DBvalue, tcase)
    Dim res, tcaseWithValues
    res = checkTest(DBvalue, expvalue, 0)
    tcaseWithValues = tcase  & vbcrlf & "期望值:" & expvalue& vbcrlf & "数据库值:" & DBvalue
    myReporter tcaseWithValues, res
End Function


'*******************************************************
'方法名:pri_verifyProperty
'功能:私有功能方法,供接口方法调用,检查测试对象的属性值与期望值是否一致,并输出测试报告
'参数:obj,测试对象 ; prop,需要检查的属性 ; expvalue,测试期望值 ; 
'     tcase,测试名称; method, 检查方法, 0代表精确匹配, 1代表模糊匹配
'*******************************************************
private Function pri_verifyProperty(obj, prop, expvalue, tcase, method)
   Dim actvalue, res, tcaseWithValues
   actvalue = obj.GetROProperty(prop)
   res = checkTest(expvalue, actvalue, method)
   tcaseWithValues = tcase & vbcrlf & "期望值:" & expvalue & vbcrlf & "实际值:" & actvalue
   MyReporter tcaseWithValues, res
End Function


'*******************************************************
'方法名:verifyProperty
'功能:检查测试对象的属性值与期望值是否一致,并输出测试报告
'参数:obj,测试对象; prop,需要检查的属性 ; expvalue,测试期望值 ;
'     tcase,测试名称; method, 检查方法, 0代表精确匹配, 1代表模糊匹配
'*******************************************************
public Function verifyProperty(obj, prop, expvalue, tcase, method)
   pri_verifyProperty obj, prop, expvalue, tcase, method 
End Function

'*******************************************************
'方法名:verifyInnerTxt
'功能:精确检查测试对象innerText属性是否符合期望值
'参数:obj,测试对象 ;expvalue,测试期望值 ; tcase,测试名称
'*******************************************************
public Function verifyInnerTxt(obj, expvalue, tcase)
	pri_verifyProperty obj, "innertext", expvalue, tcase, 0
End Function

'*******************************************************
'方法名:verifyValue
'功能:精确检查测试对象value属性是否符合期望值
'参数:obj,测试对象 ;expvalue,测试期望值 ; tcase,测试名称
'*******************************************************
public Function verifyValue(obj, expvalue, tcase)
	pri_verifyProperty obj, "value", expvalue, tcase, 0
End Function

'*******************************************************
'方法名:verifyUrl
'功能:精确检查测试对象value属性是否符合期望值
'参数:obj,测试对象 ;expvalue,测试期望值 ; tcase,测试名称
'*******************************************************
public Function verifyUrl(obj, expvalue, tcase)
	pri_verifyProperty obj, "url", expvalue, tcase, 1
End Function

'*******************************************************
'方法名:pri_verifyStyleDisplay
'功能:检查测试对象是否在页面上显示,即style的display属性是否为预期值
'参数:obj,测试对象 ; expvalue,测试期望值 ; tcase,测试名称
'*******************************************************
private Function pri_verifyStyleDisplay(obj, expvalue, tcase)
   Dim actvalue, res
   actvalue = obj.Object.currentStyle.display
   res = checkTest(expvalue, actvalue, 1)
   MyReporter tcase, res
End Function

'*******************************************************
'方法名:verifyExist
'功能:检查测试对象是否在页面上存在, 期望存在
'参数:obj,测试对象 ;  tcase,测试名称
'*******************************************************
public Function verifyExist(obj,tcase)
	If obj.exist(5) Then
		pri_verifyStyleDisplay obj, "block", tcase
		Exit Function
	End If
	MyReporter tcase, false
End Function

'*******************************************************
'方法名:verifyNotExist
'功能:检查测试对象是否在页面上存在, 期望不存在
'参数:obj,测试对象 ;  tcase,测试名称
'*******************************************************
public Function verifyNotExist(obj,tcase)
   Dim res
   res = not obj.exist(5)
   If res Then
	   MyReporter tcase, res
	   Exit Function
   Else
       pri_verifyStyleDisplay obj, "none", tcase
   End If
End Function

'*******************************************************
' 注册上面的函数
'*******************************************************

' verifyProperty
RegisterUserFunc "Browser", "verifyProperty", "verifyProperty"
RegisterUserFunc "Frame", "verifyProperty", "verifyProperty"
RegisterUserFunc "Image", "verifyProperty", "verifyProperty"
RegisterUserFunc "Link", "verifyProperty", "verifyProperty"
RegisterUserFunc "Page", "verifyProperty", "verifyProperty"
RegisterUserFunc "ViewLink", "verifyProperty", "verifyProperty"
RegisterUserFunc "WebArea", "verifyProperty", "verifyProperty"
RegisterUserFunc "WebButton", "verifyProperty", "verifyProperty"
RegisterUserFunc "WebCheckBox", "verifyProperty", "verifyProperty"
RegisterUserFunc "WebEdit", "verifyProperty", "verifyProperty"
RegisterUserFunc "WebElement", "verifyProperty", "verifyProperty"
RegisterUserFunc "WebFile", "verifyProperty", "verifyProperty"
RegisterUserFunc "WebList", "verifyProperty", "verifyProperty"
RegisterUserFunc "WebRadioGroup", "verifyProperty", "verifyProperty"
RegisterUserFunc "WebTable", "verifyProperty", "verifyProperty"
RegisterUserFunc "WebXML", "verifyProperty", "verifyProperty"

' verifyInnerTxt
RegisterUserFunc "Browser", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "Frame", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "Image", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "Link", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "Page", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "ViewLink", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "WebArea", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "WebButton", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "WebCheckBox", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "WebEdit", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "WebElement", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "WebFile", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "WebList", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "WebRadioGroup", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "WebTable", "verifyInnerTxt", "verifyInnerTxt"
RegisterUserFunc "WebXML", "verifyInnerTxt", "verifyInnerTxt"


' verifyValue
RegisterUserFunc "Browser", "verifyValue", "verifyValue"
RegisterUserFunc "Frame", "verifyValue", "verifyValue"
RegisterUserFunc "Image", "verifyValue", "verifyValue"
RegisterUserFunc "Link", "verifyValue", "verifyValue"
RegisterUserFunc "Page", "verifyValue", "verifyValue"
RegisterUserFunc "ViewLink", "verifyValue", "verifyValue"
RegisterUserFunc "WebArea", "verifyValue", "verifyValue"
RegisterUserFunc "WebButton", "verifyValue", "verifyValue"
RegisterUserFunc "WebCheckBox", "verifyValue", "verifyValue"
RegisterUserFunc "WebEdit", "verifyValue", "verifyValue"
RegisterUserFunc "WebElement", "verifyValue", "verifyValue"
RegisterUserFunc "WebFile", "verifyValue", "verifyValue"
RegisterUserFunc "WebList", "verifyValue", "verifyValue"
RegisterUserFunc "WebRadioGroup", "verifyValue", "verifyValue"
RegisterUserFunc "WebTable", "verifyValue", "verifyValue"
RegisterUserFunc "WebXML", "verifyValue", "verifyValue"

' verifyUrl
RegisterUserFunc "Frame", "verifyUrl", "verifyUrl"
RegisterUserFunc "Image", "verifyUrl", "verifyUrl"
RegisterUserFunc "Link", "verifyUrl", "verifyUrl"
RegisterUserFunc "Page", "verifyUrl", "verifyUrl"
RegisterUserFunc "WebArea", "verifyUrl", "verifyUrl"


'verifyExist
RegisterUserFunc "Browser", "verifyExist", "verifyExist"
RegisterUserFunc "Frame", "verifyExist", "verifyExist"
RegisterUserFunc "Image", "verifyExist", "verifyExist"
RegisterUserFunc "Link", "verifyExist", "verifyExist"
RegisterUserFunc "Page", "verifyExist", "verifyExist"
RegisterUserFunc "ViewLink", "verifyExist", "verifyExist"
RegisterUserFunc "WebArea", "verifyExist", "verifyExist"
RegisterUserFunc "WebButton", "verifyExist", "verifyExist"
RegisterUserFunc "WebCheckBox", "verifyExist", "verifyExist"
RegisterUserFunc "WebEdit", "verifyExist", "verifyExist"
RegisterUserFunc "WebElement", "verifyExist", "verifyExist"
RegisterUserFunc "WebFile", "verifyExist", "verifyExist"
RegisterUserFunc "WebList", "verifyExist", "verifyExist"
RegisterUserFunc "WebRadioGroup", "verifyExist", "verifyExist"
RegisterUserFunc "WebTable", "verifyExist", "verifyExist"
RegisterUserFunc "WebXML", "verifyExist", "verifyExist"

'verifyNotExist
RegisterUserFunc "Browser", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "Frame", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "Image", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "Link", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "Page", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "ViewLink", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "WebArea", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "WebButton", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "WebCheckBox", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "WebEdit", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "WebElement", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "WebFile", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "WebList", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "WebRadioGroup", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "WebTable", "verifyNotExist", "verifyNotExist"
RegisterUserFunc "WebXML", "verifyNotExist", "verifyNotExist"
