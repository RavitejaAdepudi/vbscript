Import "dynatrace.vbs"

Dim Dynatrace : Set Dynatrace = (New DynatraceHelper)("c:\temp\dynatrace_poc\version2\dtagent.dll", "VBS_INT_BACKEND", "localhost")
HandleWebRequest "index.htm"
Dynatrace.Uninitialize

Private Sub HandleWebRequest(url)
	Dynatrace.SetTag(Wscript.Arguments(0))
	Dynatrace.StartServerPP
	Dynatrace.StartPurePath "WebRequest", 10, Array(url)
		For i = 0 to 2
			ExecuteQuery("SELECT * FROM tbl WHERE id=" & i)
		Next
	Dynatrace.Leave
	Dynatrace.EndServerPP
End Sub

Private Sub ExecuteQuery(query)
	Dynatrace.Enter "ExecuteQuery", 19, Array(query)
		Wscript.Echo ("Executing query " & query)
	Dynatrace.Leave
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