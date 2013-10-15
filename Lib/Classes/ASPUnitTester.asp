<%
	Class ASPUnitTester
		Private _
			m_Responder, _
			m_Scenario

		Private _
			m_CurrentModule, _
			m_CurrentTest

		Private Sub Class_Initialize()
			Set m_Responder = New ASPUnitJSONResponder
			Set m_Scenario = New ASPUnitScenario
		End Sub

		Private Sub Class_Terminate()
			Set m_Scenario = Nothing
			Set m_Responder = Nothing
		End Sub

		Public Property Set Responder(ByRef objValue)
			Set m_Responder = objValue
		End Property

		Public Property Get Modules()
			Set Modules = m_Scenario.Modules
		End Property

		' Factory methods for private classes

		Public Function CreateModule(strName, arrTests, objLifecycle)
			Dim objReturn, _
				i

			Set objReturn = New ASPUnitModule
			objReturn.Name = strName
			For i = 0 To UBound(arrTests)
				objReturn.Tests.Add(arrTests(i))
			Next
			Set objReturn.Lifecycle = objLifecycle

			Set CreateModule = objReturn
		End Function

		Public Function CreateTest(strName)
			Dim objReturn

			Set objReturn = New ASPUnitTest
			objReturn.Name = strName

			Set CreateTest = objReturn
		End Function

		Public Function CreateLifecycle(strSetup, strTeardown)
			Dim objReturn

			Set objReturn = New ASPUnitTestLifecycle
			objReturn.Setup = strSetup
			objReturn.Teardown = strTeardown

			Set CreateLifecycle = objReturn
		End Function

		' Public methods to add modules

		Public Sub AddModule(objModule)
			Call m_Scenario.Modules.Add(objModule)
		End Sub

		Public Sub AddModules(arrModules)
			Dim i

			For i = 0 To UBound(arrModules)
				Call AddModule(arrModules(i))
			Next
		End Sub

		' Assertion Methods

		Private Function Assert(blnResult, strDescription)
			If IsObject(m_CurrentTest) Then
				m_CurrentTest.Passed = blnResult
				m_CurrentTest.Description = strDescription
			End If

			Assert = blnResult
		End Function

		Public Function Ok(blnResult, strDescription)
			Ok = Assert(blnResult, strDescription)
		End Function

		Public Function Equal(varActual, varExpected, strDescription)
			Equal = Assert((varActual = varExpected), strDescription)
		End Function

		Public Function NotEqual(varActual, varExpected, strDescription)
			NotEqual = Assert(Not (varActual = varExpected), strDescription)
		End Function

		Public Function Same(varActual, varExpected, strDescription)
			Same = Assert((varActual Is varExpected), strDescription)
		End Function

		Public Function NotSame(varActual, varExpected, strDescription)
			NotSame = Assert(Not (varActual Is varExpected), strDescription)
		End Function

		' Methods to run module tests

		Public Sub Run()
			Dim objModule, _
				i

			For i = 0 To (m_Scenario.Modules.Count - 1)
				Set objModule = m_Scenario.Modules.Item(i)
				Call RunModule(objModule)

				m_Scenario.TestCount = m_Scenario.TestCount + objModule.TestCount
				m_Scenario.PassCount = m_Scenario.PassCount + objModule.PassCount
				m_Scenario.FailCount = m_Scenario.FailCount + objModule.FailCount

				Set objModule = Nothing
			Next

			m_Responder.Respond(m_Scenario)
		End Sub

		Private Sub RunModule(ByRef objModule)
			Dim intTimeStart, _
				intTimeEnd, _
				objTest, _
				i

			Set m_CurrentModule = objModule

			intTimeStart = Timer
			For i = 0 To (objModule.Tests.Count - 1)
				Set objTest = objModule.Tests.Item(i)

				Call RunTestModuleSetup(objModule)
				Call RunModuleTest(objTest)
				Call RunTestModuleTeardown(objModule)

				If objTest.Passed Then
					objModule.PassCount = objModule.PassCount + 1
				End If

				Set objTest = Nothing
			Next
			intTimeEnd = Timer

			objModule.Duration = Round((intTimeEnd - intTimestart), 3)

			Set m_CurrentModule = Nothing
		End Sub

		Private Sub RunModuleTest(ByRef objTest)
			Set m_CurrentTest = objTest

			On Error Resume Next
			Call GetRef(objTest.Name)()

			If Err.Number <> 0 Then
				Call Assert(False, Err.Description)
			End If

			Err.Clear()
			On Error Goto 0

			Set m_CurrentTest = Nothing
		End Sub

		Private Sub RunTestModuleSetup(ByRef objModule)
			If Not objModule.Lifecycle Is Nothing Then
				If Not objModule.Lifecycle.Setup = Empty Then
					Call GetRef(objModule.Lifecycle.Setup)()
				End If
			End If
		End Sub

		Private Sub RunTestModuleTeardown(ByRef objModule)
			If Not objModule.Lifecycle Is Nothing Then
				If Not objModule.Lifecycle.Teardown = Empty Then
					Call GetRef(objModule.Lifecycle.Teardown)()
				End If
			End If
		End Sub
	End Class

	' Private Classses

	Class ASPUnitScenario
		Public _
			Modules, _
			TestCount, _
			PassCount, _
			FailCount

		Private Sub Class_Initialize()
			Set Modules = Server.CreateObject("System.Collections.ArrayList")
			PassCount = 0
			TestCount = 0
			FailCount = 0
		End Sub

		Private Sub Class_Terminate()
			Set Modules = Nothing
		End Sub

		Public Property Get Passed
			Passed = (PassCount = TestCount)
		End Property
	End Class

	Class ASPUnitModule
		Public _
			Name, _
			Tests, _
			Lifecycle, _
			Duration, _
			PassCount

		Private Sub Class_Initialize()
			Set Tests = Server.CreateObject("System.Collections.ArrayList")
			PassCount = 0
		End Sub

		Private Sub Class_Terminate()
			Set Tests = Nothing
		End Sub

		Public Property Get TestCount
			TestCount = Tests.Count()
		End Property

		Public Property Get FailCount
			FailCount = (TestCount - PassCount)
		End Property

		Public Property Get Passed
			Passed = (PassCount = TestCount)
		End Property
	End Class

	Class ASPUnitTest
		Public _
			Name, _
			Passed, _
			Description
	End Class

	Class ASPUnitTestLifecycle
		Public _
			Setup, _
			Teardown
	End Class
%>