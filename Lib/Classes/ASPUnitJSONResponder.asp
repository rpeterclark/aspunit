<%
	Class ASPUnitJSONResponder
		Private _
			m_Serializer

		Private Sub Class_Initialize()
			Set m_Serializer = New ASPUnitScenarioSerializer
		End Sub

		Private Sub Class_Terminate()
			Set m_Serializer = Nothing
		End Sub

		Public Property Set Serializer(objValue)
			Set m_Serializer = objValue
		End Property

		Public Sub Respond(objModules)
			Response.ContentType = "application/json"
			Response.Write m_Serializer.ToJSON(objModules)
		End Sub
	End Class

	' Private Classes

	Class ASPUnitScenarioSerializer
		Function ToJSON(objScenario)
			Dim objStream, _
				objModule, _
				objTest, _
				i, j, _
				strReturn

			Set objStream = Server.CreateObject("ADODB.Stream")

			objStream.Type = 2 ' adTypeText
			objStream.Mode = 3 ' adModeReadWrite
			objStream.Open

			Call objStream.WriteText("{")
			Call objStream.WriteText(JSONNumberPair("testCount", objScenario.TestCount) & ",")
			Call objStream.WriteText(JSONNumberPair("passCount", objScenario.PassCount) & ",")
			Call objStream.WriteText(JSONBooleanPair("passed", objScenario.Passed) & ",")
			Call objStream.WriteText(JSONString("modules") & ":[")
			For i = 0 To (objScenario.Modules.Count - 1)
				Set objModule = objScenario.Modules.Item(i)

				Call objStream.WriteText("{")
				Call objStream.WriteText(JSONStringPair("name", objModule.Name) & ",")
				Call objStream.WriteText(JSONNumberPair("testCount", objModule.TestCount) & ",")
				Call objStream.WriteText(JSONNumberPair("passCount", objModule.PassCount) & ",")
				Call objStream.WriteText(JSONNumberPair("failCount", (objModule.FailCount)) & ",")
				Call objStream.WriteText(JSONBooleanPair("passed", objModule.Passed) & ",")
				Call objStream.WriteText(JSONNumberPair("duration", objModule.Duration) & ",")
				Call objStream.WriteText(JSONString("tests") & ":[")
				For j = 0 To (objModule.Tests.Count - 1)
					Set objTest = objModule.Tests.Item(j)
					Call objStream.WriteText("{")
					Call objStream.WriteText(JSONStringPair("name", objTest.Name) & ",")
					Call objStream.WriteText(JSONBooleanPair("passed", objTest.Passed) & ",")
					Call objStream.WriteText(JSONStringPair("description", objTest.Description))
					Call objStream.WriteText("}")

					If j < (objModule.Tests.Count - 1) Then
						Call objStream.WriteText(",")
					End If

					Set objTest = Nothing
				Next
				Call objStream.WriteText("]")
				Call objStream.WriteText("}")

				If i < (objScenario.Modules.Count - 1) Then
					Call objStream.WriteText(",")
				End If

				Set objModule = Nothing
			Next
			Call objStream.WriteText("]")
			Call objStream.WriteText("}")

			objStream.Position = 0

			strReturn = objStream.ReadText()

			objStream.Close
			Set objStream = Nothing

			ToJSON = strReturn
		End Function

		Private Function JSONString(strValue)
			JSONString = """" & JSONStringEscape(strValue) & """"
		End Function

		Private Function JSONStringPair(strName, strValue)
			JSONStringPair = JSONString(strName) & ":" & JSONString(strValue)
		End Function

		Private Function JSONNumberPair(strName, varValue)
			JSONNumberPair = JSONString(strName) & ":" & varValue
		End Function

		Private Function JSONBooleanPair(strName, blnValue)
			JSONBooleanPair = JSONString(strName) & ":" & LCase(blnValue)
		End Function

		Private Function JSONStringEscape(strValue)
			Dim strReturn

			strReturn = strValue

			strReturn = Replace(strReturn, "\", "\\")
			strReturn = Replace(strReturn, """", "\""")
			strReturn = Replace(strReturn, vbLf, "\n")
			strReturn = Replace(strReturn, vbCr, "\n")
			strReturn = Replace(strReturn, vbCrLf, "\n")
			strReturn = Replace(strReturn, vbTab, "\t")

			JSONStringEscape = strReturn
		End Function
	End Class
%>