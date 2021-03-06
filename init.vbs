Option Explicit


' init the application
public Function initApp(program, url)
   'Close all tabs to close the browser
   If Browser(":=").Exist(3) Then
       Browser(":=").CloseAllTabs
   End If
   SystemUtil.Run program, url
End Function


''' for WebTest
' init Ie with httpwatch attachment.
' stable for x86. unstable for x64
public Function initIEWithHttpWatch(url)
   'Close all tabs to close the browser
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

' access the *url* and return the statusCode.
' plugin opened by initIEWithHttpWatch is necessary
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