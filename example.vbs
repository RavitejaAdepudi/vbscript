Import "dynatrace.vbs"
Dim Dynatrace : Set Dynatrace = (New DynatraceHelper)("c:\dtagent.dll", "VBS_INT", "localhost")

Private Sub HandleWebRequest(url)
	Dynatrace.StartPurePath "WebRequest", 5, Array(url)
		Wscript.Echo ("Handling request " & url)
		For i = 0 to 2
			ExecuteQuery("SELECT * FROM tbl WHERE id=" & i)
		Next
	Dynatrace.Leave
End Sub

Private Sub ExecuteQuery(query)
	Dynatrace.Enter "ExecuteQuery", 14, Array(query)
		Wscript.Echo ("Executing query " & query)
	Dynatrace.Leave
End Sub

HandleWebRequest "index.htm"

Dynatrace.Uninitialize

Private Sub Import(ByVal filename)
	Dim fso, sh, code, dir
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set sh  = CreateObject("WScript.Shell")
	filename = Trim(sh.ExpandEnvironmentStrings(filename))
	If Not (Left(filename, 2) = "\\" Or Mid(filename, 2, 2) = ":\") Then
		filename = fso.GetAbsolutePathName(filename)
	End If
	code = fso.OpenTextFile(filename).ReadAll
	ExecuteGlobal code
	Set fso = Nothing
	Set sh  = Nothing
End Sub
