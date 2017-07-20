Import "dynatrace.vbs"

Dim Dynatrace : Set Dynatrace = (New DynatraceHelper)("C:\Program Files\Dynatrace\Dynatrace 7.0\agent\lib\dtagent.dll", "VBS_INT", "localhost")
HandleWebRequest "index.htm"
Dynatrace.Uninitialize

Private Sub HandleWebRequest(url)
	Dynatrace.StartPurePath "WebRequest", 8, Array(url)
		For i = 0 to 2
			ExecuteQuery("SELECT * FROM tbl WHERE id=" & i)
		Next
		Dim tag
		tag = Dynatrace.GetTag()
		Dynatrace.LinkClientPurePath true, tag
		Set oShell = CreateObject ("WScript.Shell") 
		oShell.run "c:\Windows\SysWOW64\cscript.exe ""c:\temp\dynatrace_poc\version2\test - ServerSide.vbs"" " & tag
		WScript.Sleep 10000
		tag = Dynatrace.GetTag()
		Dynatrace.LinkClientPurePath false, tag
		Dynatrace.Leave
	WScript.Sleep 2000
	SyncPath tag
End Sub

Private Sub ExecuteQuery(query)
	Dynatrace.Enter "ExecuteQuery", 24, Array(query)
		Wscript.Echo ("Executing query " & query)
	Dynatrace.Leave
End Sub

Private Sub SyncPath(tag)
	Dynatrace.SetTag(tag)
	Dynatrace.StartServerPP
	Dynatrace.StartPurePath "AsyncPath", 31, Array(tag)
	WScript.Sleep 100
	Dynatrace.Leave
	Dynatrace.EndServerPP
End Sub







Private Sub Import(ByVal filename)
	Dim fso, sh, code, dir
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set sh  = CreateObject("WScript.Shell")

	filename = Trim(sh.ExpandEnvironmentStrings(filename))
	If Not (Left(filename, 2) = "\\" Or Mid(filename, 2, 2) = ":\") Then
		If Not fso.FileExists(fso.GetAbsolutePathName(filename)) Then
			For Each dir In Split(sh.ExpandEnvironmentStrings("%PATH%"), ";")
				If fso.FileExists(fso.BuildPath(dir, filename)) Then
					filename = fso.BuildPath(dir, filename)
					Exit For
				End If
			Next
		End If
		filename = fso.GetAbsolutePathName(filename)
	End If

	code = fso.OpenTextFile(filename).ReadAll

	ExecuteGlobal code

	Set fso = Nothing
	Set sh  = Nothing
End Sub