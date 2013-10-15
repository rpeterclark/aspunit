<% Option Explicit %>

<!-- #include file="../Lib/ASPUnit.asp" -->

<%
	Dim objLifecycle
	Set objLifecycle = ASPUnit.CreateLifeCycle("Setup", "Teardown")

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester Run Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterRunsSetup"), _
				ASPUnit.CreateTest("ASPUnitTesterRunsTeardown"), _
				ASPUnit.CreateTest("ASPUnitTesterRunsTests"), _
				ASPUnit.CreateTest("ASPUnitTesterPassesToResponder") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester Result Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterTestResultTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterTestResultFalsey"), _
				ASPUnit.CreateTest("ASPUnitTesterModulePassCount"), _
				ASPUnit.CreateTest("ASPUnitTesterModuleFailCount"), _
				ASPUnit.CreateTest("ASPUnitTesterModulePassedTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterModulePassedFalsey"), _
				ASPUnit.CreateTest("ASPUnitTesterScenarioTestCount"), _
				ASPUnit.CreateTest("ASPUnitTesterScenarioPassCount"), _
				ASPUnit.CreateTest("ASPUnitTesterScenarioFailCount") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.Run()

	' Create a global instance of ASPUnitTester for testing

	Sub Setup()
		Call ExecuteGlobal("Dim objService")
		Set objService = New ASPUnitTester
	End Sub

	Sub Teardown()
		Set objService = Nothing
	End Sub

	' Create mock responder service to prevent output

	Class ASPUnitTesterMockResponder
		Public _
			Scenario

		Private Sub Class_Initialize()
			Set Scenario = Nothing
		End Sub

		Private Sub Class_Terminate()
			Set Scenario = Nothing
		End Sub

		Public Sub Respond(objValue)
			Set Scenario = objValue
		End Sub
	End Class

	' Mock lifecycle methods that set global variables to indicate they executed

	Sub MockSetup()
		Call ExecuteGlobal("Dim blnSetupRan")
		blnSetupRan = True
	End Sub

	Sub MockTeardown()
		Call ExecuteGlobal("Dim blnTeardownRan")
		blnTeardownRan = True
	End Sub

	' Mock tests to be executed for various situations

	Sub MockTestPass()
		Call objService.Ok(True, "Mock test")
	End Sub

	Sub MockTestFail()
		Call objService.Ok(False, "Mock test")
	End Sub

	Sub MockTestFirst()
		Call ExecuteGlobal("Dim intTestTotal")
		intTestTotal = 1

		Call objService.Ok(True, "Mock test")
	End Sub

	Sub MockTestMiddle()
		intTestTotal = intTestTotal + 1

		Call objService.Ok(True, "Mock test")
	End Sub

	Sub MockTestLast()
		intTestTotal = intTestTotal + 1

		Call objService.Ok(True, "Mock test")
	End Sub

	' Test that ASPUnitTester runs lifecycle setup method if specified

	Sub ASPUnitTesterRunsSetup()
		Set objService.Responder = New ASPUnitTesterMockResponder

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestPass") _
				), _
				objService.CreateLifecycle( _
					"MockSetup", _
					"" _
				) _
			) _
		)

		objService.Run()

		Call ASPUnit.Equal(blnSetupRan, True, "Run should execute specified setup method")
	End Sub

	' Test that ASPUnitTester runs lifecycle teardown method if specified

	Sub ASPUnitTesterRunsTeardown()
		Set objService.Responder = New ASPUnitTesterMockResponder

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestPass") _
				), _
				objService.CreateLifecycle( _
					"", _
					"MockTeardown" _
				) _
			) _
		)

		objService.Run()

		Call ASPUnit.Equal(blnTeardownRan, True, "Run should execute specified teardown method")
	End Sub

	' Test that ASPUnitTester runs all specified tests

	Sub ASPUnitTesterRunsTests()
		Set objService.Responder = New ASPUnitTesterMockResponder

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestFirst"), _
					objService.CreateTest("MockTestMiddle"), _
					objService.CreateTest("MockTestLast") _
				), _
				Nothing _
			) _
		)

		objService.Run()

		Call ASPUnit.Equal(intTestTotal, 3, "Run should execute specified test methods")
	End Sub

	' Test that ASPUnitTester implements responder service

	Sub ASPUnitTesterPassesToResponder()
		Dim objResponder
		Set objResponder = New ASPUnitTesterMockResponder
		Set objService.Responder = objResponder

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestPass") _
				), _
				Nothing _
			) _
		)

		objService.Run()

		Dim objScenario
		Set objScenario = objResponder.Scenario

		Call ASPUnit.NotEqual(TypeName(objScenario), "Nothing", "Run should mark passed tests as passed")

		Set objScenario = Nothing
		Set objResponder = Nothing
	End Sub

	' Test that ASPUnitTester flags test results appropriately

	Sub ASPUnitTesterTestResultTruthy()
		Dim objResponder
		Set objResponder = New ASPUnitTesterMockResponder
		Set objService.Responder = objResponder

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestPass") _
				), _
				Nothing _
			) _
		)

		objService.Run()

		Dim objScenario
		Set objScenario = objResponder.Scenario

		Call ASPUnit.Equal(objScenario.Modules.Item(0).Tests.Item(0).Passed, True, "Run should mark passed tests as passed")

		Set objScenario = Nothing
		Set objResponder = Nothing
	End Sub

	Sub ASPUnitTesterTestResultFalsey()
		Dim objResponder
		Set objResponder = New ASPUnitTesterMockResponder
		Set objService.Responder = objResponder

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestFail") _
				), _
				Nothing _
			) _
		)

		objService.Run()

		Dim objScenario
		Set objScenario = objResponder.Scenario

		Call ASPUnit.Equal(objScenario.Modules.Item(0).Tests.Item(0).Passed, False, "Run should mark failed tests as failed")

		Set objScenario = Nothing
		Set objResponder = Nothing
	End Sub

	' Test that ASPUnitTester correctly aggregates pass/fail totals at Module level

	Sub ASPUnitTesterModulePassCount()
		Dim objResponder
		Set objResponder = New ASPUnitTesterMockResponder
		Set objService.Responder = objResponder

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestPass"), _
					objService.CreateTest("MockTestPass"), _
					objService.CreateTest("MockTestFail") _
				), _
				Nothing _
			) _
		)

		objService.Run()

		Dim objScenario
		Set objScenario = objResponder.Scenario

		Call ASPUnit.Equal(objScenario.Modules.Item(0).PassCount, 2, "Run should count passed tests")

		Set objScenario = Nothing
		Set objResponder = Nothing
	End Sub

	Sub ASPUnitTesterModuleFailCount()
		Dim objResponder
		Set objResponder = New ASPUnitTesterMockResponder
		Set objService.Responder = objResponder

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestPass"), _
					objService.CreateTest("MockTestFail"), _
					objService.CreateTest("MockTestFail") _
				), _
				Nothing _
			) _
		)

		objService.Run()

		Dim objScenario
		Set objScenario = objResponder.Scenario

		Call ASPUnit.Equal(objScenario.Modules.Item(0).FailCount, 2, "Run should count failed tests")

		Set objScenario = Nothing
		Set objResponder = Nothing
	End Sub

	' Test that ASPUnitTester flags module results appropriately

	Sub ASPUnitTesterModulePassedTruthy()
		Dim objResponder
		Set objResponder = New ASPUnitTesterMockResponder
		Set objService.Responder = objResponder

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestPass"), _
					objService.CreateTest("MockTestPass") _
				), _
				Nothing _
			) _
		)

		objService.Run()

		Dim objScenario
		Set objScenario = objResponder.Scenario

		Call ASPUnit.Equal(objScenario.Modules.Item(0).Passed, True, "Run should mark module with all passing tests as passed")

		Set objScenario = Nothing
		Set objResponder = Nothing
	End Sub

	Sub ASPUnitTesterModulePassedFalsey()
		Dim objResponder
		Set objResponder = New ASPUnitTesterMockResponder
		Set objService.Responder = objResponder

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestFail"), _
					objService.CreateTest("MockTestPass") _
				), _
				Nothing _
			) _
		)

		objService.Run()

		Dim objScenario
		Set objScenario = objResponder.Scenario

		Call ASPUnit.Equal(objScenario.Modules.Item(0).Passed, False, "Run should mark module with any failing tests as failed")

		Set objScenario = Nothing
		Set objResponder = Nothing
	End Sub

	' Test that ASPUnitTester correctly aggregates totals at Scenario level

	Sub ASPUnitTesterScenarioTestCount()
		Dim objResponder
		Set objResponder = New ASPUnitTesterMockResponder
		Set objService.Responder = objResponder

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestPass") _
				), _
				Nothing _
			) _
		)

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestPass") _
				), _
				Nothing _
			) _
		)

		objService.Run()

		Dim objScenario
		Set objScenario = objResponder.Scenario

		Call ASPUnit.Equal(objScenario.TestCount, 2, "Run should aggregate module test count")

		Set objScenario = Nothing
		Set objResponder = Nothing
	End Sub

	Sub ASPUnitTesterScenarioPassCount()
		Dim objResponder
		Set objResponder = New ASPUnitTesterMockResponder
		Set objService.Responder = objResponder

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestPass") _
				), _
				Nothing _
			) _
		)

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestPass") _
				), _
				Nothing _
			) _
		)

		objService.Run()

		Dim objScenario
		Set objScenario = objResponder.Scenario

		Call ASPUnit.Equal(objScenario.PassCount, 2, "Run should aggregate module pass count")

		Set objScenario = Nothing
		Set objResponder = Nothing
	End Sub

	Sub ASPUnitTesterScenarioFailCount()
		Dim objResponder
		Set objResponder = New ASPUnitTesterMockResponder
		Set objService.Responder = objResponder

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestFail") _
				), _
				Nothing _
			) _
		)

		Call objService.AddModule( _
			objService.CreateModule( _
				"", _
				Array( _
					objService.CreateTest("MockTestFail") _
				), _
				Nothing _
			) _
		)

		objService.Run()

		Dim objScenario
		Set objScenario = objResponder.Scenario

		Call ASPUnit.Equal(objScenario.FailCount, 2, "Run should aggregate module fail count")

		Set objScenario = Nothing
		Set objResponder = Nothing
	End Sub
%>