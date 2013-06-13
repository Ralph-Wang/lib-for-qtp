Option Explicit


'****************************************************************************************************************************
'函数名:IdentifyDOM(htmlid)
'功能:通过htmlid识别DOM对象; 返回对象为对应DOM对象
'参数:htmlid为需要识别的DOM对象的id属性
'****************************************************************************************************************************
Function IdentifyDOM(htmlid)
    'MsgBox htmlid
	Set IdentifyDOM = Browser(":=").Page(":=").Object.GetElementById(htmlid)
End Function

'****************************************************************************************************************************
'函数名:IdentifyDOMfromIframe(iframe,htmlid)
'功能:从iframe分页中,通过htmlid识别DOM对象; 返回对象为对应DOM对象
'参数:iframe为容器iframe的DOM对象,htmlid为需要识别的DOM对象的id属性
'****************************************************************************************************************************
Function IdentifyDOMfromIframe(iframe,htmlid)
	If not isObject(iframe) Then
		msgbox "IdentifyDOMfromIframe函数的iframe参数应该是html对象"
		Exit Function
	End If
	Set IdentifyDOMfromIframe = iframe.contentDocument.getElementById(htmlid)
End Function

'****************************************************************************************************************************
'函数名:SetInputValue
'功能:对input标签输入值
'参数:htmlid, 输入对象的htmlid属性
'****************************************************************************************************************************
Function SetInputValue(htmlid,s_value)
	Dim obj
    Set obj = IdentifyDOM(htmlid)
    obj.value = s_value
End Function