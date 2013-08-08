Option Explicit


'****************************************************************************************************************************
'函数名:getDomByID(htmlid)
'功能:通过htmlid识别DOM对象; 返回对象为对应DOM对象
'参数:htmlid为需要识别的DOM对象的id属性
'****************************************************************************************************************************
Function getDomByID(htmlid)
    'MsgBox htmlid
	Set getDomByID = Browser(":=").Page(":=").Object.GetElementById(htmlid)
End Function

'****************************************************************************************************************************
'函数名:getDomByClassName(className)
'功能:通过className识别DOM对象; 返回对象为对应DOM对象(若同一页面有多个同class标签,则返回第一个)
'参数:className为需要识别的DOM对象的class属性
'****************************************************************************************************************************
Function getDomByClassName(className)
    'MsgBox className
	Set getDomByClassName = Browser(":=").Page(":=").Object.GetElementsByClassName(className)(0)
End Function

'****************************************************************************************************************************
'函数名:getDomByName(name)
'功能:通过name识别DOM对象; 返回对象为对应DOM对象(若同一页面有多个同name标签,则返回第一个)
'参数:className为需要识别的DOM对象的name属性
'****************************************************************************************************************************
Function getDomByName(name)
    'MsgBox className
	Set getDomByName = Browser(":=").Page(":=").Object.GetElementsByName(name)(0)
End Function

'****************************************************************************************************************************
'函数名:getDomfromIframe(iframe,htmlid)
'功能:从iframe分页中,通过htmlid识别DOM对象; 返回对象为对应DOM对象
'参数:iframe为容器iframe的DOM对象,htmlid为需要识别的DOM对象的id属性
'****************************************************************************************************************************
Function getDomfromIframe(iframe,htmlid)
	If not isObject(iframe) Then
		msgbox "getDomfromIframe函数的iframe参数应该是html对象"
		Exit Function
	End If
	Set getDomfromIframe = iframe.contentDocument.getElementById(htmlid)
End Function

'****************************************************************************************************************************
'函数名:getDomfromIframeByClassName(iframe,className)
'功能:从iframe分页中,通过className识别DOM对象; 返回对象为对应DOM对象
'参数:iframe为容器iframe的DOM对象,className为需要识别的DOM对象的class属性
'****************************************************************************************************************************
Function getDomfromIframeByName(iframe,className)
	If not isObject(iframe) Then
		msgbox "getDomfromIframeByClassName函数的iframe参数应该是html对象"
		Exit Function
	End If
	Set getDomfromIframeByName = iframe.contentDocument.getElementsByClassName(className)(0)
End Function

'****************************************************************************************************************************
'函数名:getDomfromIframeByName(iframe,name)
'功能:从iframe分页中,通过name识别DOM对象; 返回对象为对应DOM对象
'参数:iframe为容器iframe的DOM对象,className为需要识别的DOM对象的name属性
'****************************************************************************************************************************
Function getDomfromIframeByName(iframe,name)
	If not isObject(iframe) Then
		msgbox "getDomfromIframeByName函数的iframe参数应该是html对象"
		Exit Function
	End If
	Set getDomfromIframeByName = iframe.contentDocument.getElementsByClassName(name)(0)
End Function
