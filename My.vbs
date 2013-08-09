Option Explicit

''' 与httpWatch结合的函数
'****************************************************************************************************************************
'函数名:initIEWithHttpWatch
'功能:重新启动浏览器,并访问被测系统; HttpWatch的plugin对象
'参数:url,被测系统的网页地址
'****************************************************************************************************************************
public Function initIEWithHttpWatch(url)
   '用描述性编程关闭 浏览器对象 的所有 标签
   If Browser(":=").Exist(3) Then
       Browser(":=").CloseAllTabs
   End If
   Dim ct, plugin, ie
   Set ct = CreateObject("HttpWatch.Controller")
   Set ie = CreateObject("internetexplorer.application")
   ie.Visible  = True
   Set plugin = ct.IE.Attach(ie)
   plugin.Log.EnableFilter False
   plugin.GotoURL url
   Set initIEWithHttpWatch = plugin
   Set plugin = Nothing
   Set ct = Nothing
   Set ie = Nothing
End Function

'****************************************************************************************************************************
'函数名:getStatusCode
'功能:访问url,并返回StatusCode
'参数: plugin,HttpWatch的plugin类,url待访问地址
'****************************************************************************************************************************
public Function getStatusCode(plugin, url)
	Dim ct
	Set ct = CreateObject("HttpWatch.Controller")
	plugin.Clear
	plugin.Record
	plugin.GotoURL url
	ct.Wait plugin, -1
	plugin.stop
	getStatusCode = plugin.log.Entries.item(0).StatusCode
End Function

'' 一般的函数
'****************************************************************************************************************************
'函数名:initApp
'功能:重新启动浏览器,并访问被测系统; 无返回值
'参数:program,测试机IE浏览器的绝对路径  ; url,被测系统的网页地址
'****************************************************************************************************************************
public Function initApp(program, url)
   '用描述性编程关闭 浏览器对象 的所有 标签
   If Browser(":=").Exist(3) Then
       Browser(":=").CloseAllTabs
   End If
   SystemUtil.Run program, url
End Function

'****************************************************************************************************************************
'函数名:randInt
'功能:随机生成一个整数 (0 ~ num)
'参数: num随机数的最大值
'****************************************************************************************************************************
public Function randInt(num)
    randInt = RandomNumber.Value(0, num)
End Function

'****************************************************************************************************************************
'函数名:randFloat
'功能:随机生成一个浮点数(0 ~ num)
'参数: num随机数的最大值,小数位数dotNum
'****************************************************************************************************************************
public Function randFloat(num, dotNum)
   Dim intger
   dotNum = 10 ^ dotNum
   intger = RandomNumber.Value(0, num * dotNum)
   randFloat = intger / dotNum
End Function
'****************************************************************************************************************************
'函数名:randStr
'功能:随机生成一定长度的字符串, 大小写混合
'参数: num表示生成长度
'****************************************************************************************************************************
public Function randStr(num)
    Dim i, str
    For i = 1 To num Step 1
        If RandomNumber.Value(0,1) = 0 Then
            str = str & chr(RandomNumber.Value(65,90))
        Else
            str = str & chr(RandomNumber.Value(97,122))
        End If     
    Next
    randStr = str
End Function


'*****************************************************************************
'以下部分为测试对象的注册方法,用以检查测试结果及输出测试报告
'*****************************************************************************

'****************************************************************************************************************************
'函数名:checkTest
'功能:检查测试实际值与用例期望值是否一致; 一致则返回True,否则返回False
'参数:expvalue,用例期望值   ;  actvalue, 测试实际值   ; method, 0代表精确匹配, 1代表模糊匹配
'****************************************************************************************************************************
private Function checkTest(expvalue, actvalue, method)
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

'****************************************************************************************************************************
'函数名:myReporter
'功能:根据测试结果输出标准测试报告; 无返回值
'参数:tcase,用例名称 ; res,布尔值,测试结果
'****************************************************************************************************************************
private Function myReporter(tcaseWithValues, res)
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

'****************************************************************************************************************************
'函数名:verifyDBres
'功能:根据测试结果输出标准测试报告; 无返回值
'参数:tcase,用例名称 ; res,布尔值,测试结果
'****************************************************************************************************************************
public Function verifyDBvalue(expvalue, DBvalue, tcase)
    Dim res, tcaseWithValues
    res = checkTest(DBvalue, expvalue, 0)
    tcaseWithValues = tcase  & vbcrlf & "期望值:" & expvalue& vbcrlf & "数据库值:" & DBvalue
    myReporter tcaseWithValues, res
End Function


'****************************************************************************************************************************
'方法名:pri_verifyProperty
'功能:私有功能方法,供接口方法调用,检查测试对象的属性值与期望值是否一致,并输出测试报告
'参数:obj,测试对象 ; prop,需要检查的属性 ; expvalue,测试期望值 ; tcase,测试名称; method, 检查方法, 0代表精确匹配, 1代表模糊匹配
'****************************************************************************************************************************
private Function pri_verifyProperty(obj, prop, expvalue, tcase, method)
   Dim actvalue, res, tcaseWithValues
   actvalue = obj.GetROProperty(prop)
   res = checkTest(expvalue, actvalue, method)
   tcaseWithValues = tcase & vbcrlf & "期望值:" & expvalue & vbcrlf & "实际值:" & actvalue
   MyReporter tcaseWithValues, res
End Function


'****************************************************************************************************************************
'方法名:verifyProperty
'功能:检查测试对象的属性值与期望值是否一致,并输出测试报告
'参数:obj,测试对象(调用时默认传入); prop,需要检查的属性 ; expvalue,测试期望值 ; tcase,测试名称; method, 检查方法, 0代表精确匹配, 1代表模糊匹配
'****************************************************************************************************************************
public Function verifyProperty(obj, prop, expvalue, tcase, method)
   pri_verifyProperty obj, prop, expvalue, tcase, method 
End Function

'****************************************************************************************************************************
'方法名:verifyInnerTxt
'功能:精确检查测试对象innerText属性是否符合期望值
'参数:obj,测试对象 ;expvalue,测试期望值 ; tcase,测试名称
'****************************************************************************************************************************
public Function verifyInnerTxt(obj, expvalue, tcase)
	pri_verifyProperty obj, "innertext", expvalue, tcase, 0
End Function

'****************************************************************************************************************************
'方法名:verifyValue
'功能:精确检查测试对象value属性是否符合期望值
'参数:obj,测试对象 ;expvalue,测试期望值 ; tcase,测试名称
'****************************************************************************************************************************
public Function verifyValue(obj, expvalue, tcase)
	pri_verifyProperty obj, "value", expvalue, tcase, 0
End Function

'****************************************************************************************************************************
'方法名:verifyUrl
'功能:精确检查测试对象value属性是否符合期望值
'参数:obj,测试对象 ;expvalue,测试期望值 ; tcase,测试名称
'****************************************************************************************************************************
public Function verifyUrl(obj, expvalue, tcase)
	pri_verifyProperty obj, "url", expvalue, tcase, 1
End Function




'****************************************************************************************************************************
'方法名:pri_verifyStyleDisplay
'功能:检查测试对象是否在页面上显示,即style的display属性是否为预期值
'参数:obj,测试对象 ; prop,需要检查的属性 ; expvalue,测试期望值 ; tcase,测试名称; method, 检查方法, 0代表精确匹配, 1代表模糊匹配
'****************************************************************************************************************************
private Function pri_verifyStyleDisplay(obj, expvalue, tcase)
   Dim actvalue, res
   actvalue = obj.Object.currentStyle.display
   res = checkTest(expvalue, actvalue, 1)
   MyReporter tcase, res
End Function

'****************************************************************************************************************************
'方法名:verifyExist
'功能:检查测试对象是否在页面上存在, 期望存在
'参数:obj,测试对象 ;  tcase,测试名称
'****************************************************************************************************************************
public Function verifyExist(obj,tcase)
	If obj.exist(5) Then
		pri_verifyStyleDisplay obj, "block", tcase
		Exit Function
	End If
	MyReporter tcase, false
End Function

'****************************************************************************************************************************
'方法名:verifyNotExist
'功能:检查测试对象是否在页面上存在, 期望不存在
'参数:obj,测试对象 ;  tcase,测试名称
'****************************************************************************************************************************
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



'****************************************************************************************************************************
' 注册上面的函数
'****************************************************************************************************************************

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
