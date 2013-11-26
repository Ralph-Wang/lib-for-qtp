'''''''''''''''''
' > init.vbs
'''''''''''''''''
initApp "C:\Program Files\Internet Explorer\iexplore.exe", "about:blank"

assert Browser(":=").Exist(5), "没有启动浏览器"
assert Browser(":=").Page(":=").GetROProperty("url")="about:blan1k", "启动页不是about:blank"

Set plugin = initIEWithHttpWatch("about:blank")

assert Browser(":=").Exist(5), "没有启动浏览器"
assert Browser(":=").Page(":=").GetROProperty("url")="about:blank", "启动页不是about:blank"

statusCode = getStatusCode(plugin, "http://www.baidu.com")

assert statusCode >= 200, "获取http状态码失败"

'''''''''''''''''
' > Dom.vbs
'''''''''''''''''
initApp "C:\Program Files\Internet Explorer\iexplore.exe", "http://115.29.162.102/qtplib.html"
Set btn1 = getDomByID("btn")
assert btn1.value = "btn", "没有找到ID为btn1的元素"

Set btn3 = getDomByName("btn")
assert btn3.value = "btn", "没有找到name为btn的元素"


'''''''''''''''''
' > DataGen.vbs
'''''''''''''''''

theSelected = Browser(":=").Page(":=").WebList("html id:=theSelect").randomSelect()
assert (inStr(1, "1234", theSelected) > 0), "随机选择应该在可选范围内(1234)"

oStr = "1234"
res = zFill(oStr, 3)
assert (oStr = res), "num < 原字符串长时，应返回原值"
res = zFill(oStr, 4)
assert (oStr = res), "num = 原字符串长时，应返回原值"
res = zFill(oStr, 5)
assert (("0" + oStr) = res), "应补1位0"


theRandInt = randInt(5, 20)
assert ((theRandInt >= 5) and (theRandInt <=20)), "随机数应该在5~20之间"


theRandFloat = randFloat(5, 20, 3)
assert ((theRandFloat >= 5) and (theRandFloat <=20)), "随机数应该在5~20之间"
assert (len(split(cStr(theRandFloat),".")(1)) <= 3), "小数点后应该小于等于3位"


theRandAlpha = randAlpha(8)
assert (len(theRandAlpha) = 8), "应该生成8位字符串"

theRandUnique = randUnique()
assert (len(theRandUnique) = 14), "应该生成14位唯一字符串"


'''''''''''''''''
' > ReadConfigFromExcel.vbs
'''''''''''''''''
' unstable for x86
'Set theDict = GetConfigFromExcel("C:\Users\Administrator\Desktop\github\lib-for-qtp\test.xls","test","sKey","sValue")
'assert (theDict.Item("theItem") = "theValue"), "theItem对应的值应该是theValue"
'' an implementation of assertion

'''''''''''''''''
' > CheckPoint.vbs
'''''''''''''''''
theMatch = regSearch("\d+", "test123match")
assert theMatch="123", "没有找都匹配的123"

theRes = assertEqual(1, 1, 0)
assert theRes=true, "1=1断言应为真"
theRes = assertEqual(1, 2, 0)
assert theRes=false, "1=2断言应为假"

theRes = assertEqual("一部分", "一部分123!!!", 1)
assert theRes=true, "'一部分', '一部分123!!!'模糊匹配断言应为真"
theRes = assertEqual("一部分123!!!", "一部分", 1)
assert theRes=true, "'一部分123!!!', '一部分' 模糊匹配断言应为真"
theRes = assertEqual("一部分123!!!", "xyz", 1)
assert theRes=false, "'一部分123!!!', 'xyz' 模糊匹配断言应为假"

Function assert(expression, errDescription)
   If expression Then
	   reporter.ReportEvent micPass, "Assertion", "Pass"
   Else
       reporter.ReportEvent micFail, "Assertion", errDescription
   End If
End Function

























