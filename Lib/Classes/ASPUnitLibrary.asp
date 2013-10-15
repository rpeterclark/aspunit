<%
	Class ASPUnitLibrary
		Private _
			m_Runner, _
			m_Tester, _
			m_Task

		Public _
			Version

		Private Sub Class_Initialize()
			Version = "0.1.0"

			Set m_Runner = New ASPUnitRunner
			Set m_Tester = New ASPUnitTester
		End Sub

		Private Sub Class_Terminate()
			Set m_Runner = Nothing
			Set m_Tester = Nothing
		End Sub

		Public Property Set Runner(ByRef objValue)
			Set m_Runner = objValue
		End Property

		Public Property Set Tester(ByRef objValue)
			Set m_Tester = objValue
		End Property

		Public Property Let Task(strTask)
			m_Task = strTask
		End Property

		Public Sub Run()
			Select Case UCase(m_Task)
				Case "TEST"
					Call m_Tester.Run()
				Case Else
					Call m_Runner.Run()
			End Select
		End Sub

		' Test Service Facade

		Public Property Set Responder(ByRef objValue)
			Set m_Tester.Responder = objValue
		End Property

		Public Function CreateModule(strName, arrTests, objLifecycle)
			Set CreateModule = m_Tester.CreateModule(strName, arrTests, objLifecycle)
		End Function

		Public Function CreateTest(strName)
			Set CreateTest = m_Tester.CreateTest(strName)
		End Function

		Public Function CreateLifecycle(strSetup, strTeardown)
			Set CreateLifecycle = m_Tester.CreateLifecycle(strSetup, strTeardown)
		End Function

		Public Sub AddModule(objModule)
			Call m_Tester.AddModule(objModule)
		End Sub

		Public Sub AddModules(arrModules)
			Call m_Tester.AddModules(arrModules)
		End Sub

		Public Function Ok(blnResult, strDescription)
			Call m_Tester.Ok(blnResult, strDescription)
		End Function

		Public Function Equal(varActual, varExpected, strDescription)
			Call m_Tester.Equal(varActual, varExpected, strDescription)
		End Function

		Public Function NotEqual(varActual, varExpected, strDescription)
			Call m_Tester.NotEqual(varActual, varExpected, strDescription)
		End Function

		Public Function Same(varActual, varExpected, strDescription)
			Call m_Tester.Same(varActual, varExpected, strDescription)
		End Function

		Public Function NotSame(varActual, varExpected, strDescription)
			Call m_Tester.NotSame(varActual, varExpected, strDescription)
		End Function

		' UI Service Facade

		Public Property Set Theme(ByRef objValue)
			Set m_Runner.Theme = objValue
		End Property

		Public Sub AddPage(strPage)
			Call m_Runner.AddPage(strPage)
		End Sub

		Public Sub AddPages(arrPages)
			Call m_Runner.AddPages(arrPages)
		End Sub
	End Class
%>