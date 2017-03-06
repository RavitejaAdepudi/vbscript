dim dynatraceEnteredMethodId()
dim dynatraceEnteredSerialNo()
dim dynatraceADKObj

Class DynatraceHelper

	Public Default Function Init(agentLibrary, agentName, server)
		If dynatraceADKObj = Empty Then
			Set dynatraceADKObj = createobject("Dynatrace.ADK")
			dynatraceADKObj.set "agentlibrary=" & agentLibrary
			dynatraceADKObj.set "agentname=" & agentName
			dynatraceADKObj.set "server=" & server
			dynatraceADKObj.initialize 0, ""
			ReDim Preserve dynatraceEnteredMethodId(0)
			ReDim Preserve dynatraceEnteredSerialNo(0)
			Set Init = Me
		End If
	End Function

	Public Function StartPurePath(methodName, line, arguments)
		EnterInternal methodName, line, 1, arguments
	End Function

	Public Function Enter(methodName, line, arguments)
		EnterInternal methodName, line, 0, arguments
	End Function

	Private Function EnterInternal(methodName, line, startPurePath, arguments)
		dim dt_method_id
		dim scriptName
		if IsObject(Wscript) then
			scriptName = Wscript.ScriptFullName
		else
			scriptName = Request.ServerVariables("SCRIPT_NAME")
		end if
		
		dt_method_id = dynatraceADKObj.get_method_id(methodName, scriptName, line, "VBS", "0")

		dim dt_serial_no
		dt_serial_no = dynatraceADKObj.get_serial_no(dt_method_id, startPurePath)

		ReDim Preserve dynatraceEnteredMethodId(UBound(dynatraceEnteredMethodId) + 1)
		ReDim Preserve dynatraceEnteredSerialNo(UBound(dynatraceEnteredSerialNo) + 1)
		dynatraceEnteredMethodId(UBound(dynatraceEnteredMethodId)) = UBound(dynatraceEnteredMethodId)
		dynatraceEnteredSerialNo(UBound(dynatraceEnteredSerialNo)) = UBound(dynatraceEnteredSerialNo)
		For Each argument In arguments
			dynatraceADKObj.capture_string dt_serial_no, argument
		Next
		dynatraceEnteredMethodId(UBound(dynatraceEnteredMethodId)) = dt_method_id
		dynatraceEnteredSerialNo(UBound(dynatraceEnteredSerialNo)) = dt_serial_no
		dynatraceADKObj.enter dt_method_id, dt_serial_no
	End Function

	Public Function Leave()
		dynatraceADKObj.exit dynatraceEnteredMethodId(UBound(dynatraceEnteredMethodId)), dynatraceEnteredSerialNo(UBound(dynatraceEnteredSerialNo))
		ReDim Preserve dynatraceEnteredMethodId(UBound(dynatraceEnteredMethodId) - 1)
		ReDim Preserve dynatraceEnteredSerialNo(UBound(dynatraceEnteredSerialNo) - 1)
	End Function

	Public Function GetTag()
		GetTag = dynatraceADKObj.get_tag_as_string
	End Function

	Public Function SetTag(tag)
		dynatraceADKObj.set_tag_from_string(tag)
	End Function

	Public Function LinkClientPurePath(synchronous, tag)
		LinkClientPurePath = dynatraceADKObj.link_client_purepath_by_string(synchronous, tag)
	End Function

	Public Function Uninitialize()
		dynatraceADKObj.uninitialize
		dynatraceADKObj = Empty
	End Function
 End Class