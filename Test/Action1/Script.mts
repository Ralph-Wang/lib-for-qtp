' init.vbs
initApp "C:\Program Files\Internet Explorer\iexplore.exe", "about:blank"

assert Browser(":=").Exist(5), "没有启动浏览器"
assert Browser(":=").Page(":=").GetROProperty("url")="about:blan1k", "启动页不是about:blank"

Set plugin = initIEWithHttpWatch("about:blank")

assert Browser(":=").Exist(5), "没有启动浏览器"
assert Browser(":=").Page(":=").GetROProperty("url")="about:blank", "启动页不是about:blank"

statusCode = getStatusCode(plugin, "http://www.baidu.com")

assert statusCode >= 200, "获取http状态码失败"


'' an implementation of assertion
Function assert(expression, errDescription)
   If expression Then
	   reporter.ReportEvent micPass, "Assertion", "Pass"
   Else
       reporter.ReportEvent micFail, "Assertion", errDescription
   End If
End Function






