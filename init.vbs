Option Explicit

''' 与httpWatch结合的函数
'*******************************************************
'函数名:initIEWithHttpWatch
'功能:重新启动浏览器,并访问被测系统; HttpWatch的plugin对象
'参数:url,被测系统的网页地址
'*******************************************************
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

'*******************************************************
'函数名:getStatusCode
'功能:访问url,并返回StatusCode
'参数: plugin,HttpWatch的plugin类,url待访问地址
'*******************************************************
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

'*******************************************************
'函数名:initApp
'功能:重新启动浏览器,并访问被测系统; 无返回值
'参数:program,测试机IE浏览器的绝对路径 ;url,被测系统的网页地址
'*******************************************************
public Function initApp(program, url)
   '用描述性编程关闭 浏览器对象 的所有 标签
   If Browser(":=").Exist(3) Then
       Browser(":=").CloseAllTabs
   End If
   SystemUtil.Run program, url
End Function
