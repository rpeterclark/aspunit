<% Option Explicit %>

<!-- #include file="../Lib/ASPUnit.asp" -->

<%
	Dim objLifecycle
	Set objLifecycle = ASPUnit.CreateLifeCycle("Setup", "Teardown")

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester CreateModule Factory Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterCreateModule"), _
				ASPUnit.CreateTest("ASPUnitTesterCreateModuleName"), _
				ASPUnit.CreateTest("ASPUnitTesterCreateModuleTests"), _
				ASPUnit.CreateTest("ASPUnitTesterCreateModuleLifecycle") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester CreateTest Factory Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterCreateTest"), _
				ASPUnit.CreateTest("ASPUnitTesterCreateTestName") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester CreateLifecyle Factory Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterCreateLifecycle"), _
				ASPUnit.CreateTest("ASPUnitTesterCreateLifecycleSetup"), _
				ASPUnit.CreateTest("ASPUnitTesterCreateLifecycleTeardown") _
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

	' Test that ASPUnitTester CreateModule return expected objects

	Sub ASPUnitTesterCreateModule()
		Dim objModule

		Set objModule = objService.CreateModule("Test", Array(), Nothing)
		Call ASPUnit.Equal(TypeName(objModule), "ASPUnitModule", "CreateModule factory method should return ASPUnitModule")
		Set objModule = Nothing
	End Sub

	Sub ASPUnitTesterCreateModuleName()
		Dim objModule

		Set objModule = objService.CreateModule("Test", Array(), Nothing)
		Call ASPUnit.Equal(objModule.Name, "Test", "CreateModule factory method should return ASPUnitModule with specified name")
		Set objModule = Nothing
	End Sub

	Sub ASPUnitTesterCreateModuleTests()
		Dim objModule

		Set objModule = objService.CreateModule("Test", Array(Nothing, Nothing), Nothing)
		Call ASPUnit.Equal(objModule.Tests.Count, 2, "CreateModule factory method should return ASPUnitModule with tests")
		Set objModule = Nothing
	End Sub

	Sub ASPUnitTesterCreateModuleLifecycle()
		Dim objModule

		Set objModule = objService.CreateModule("Test", Array(), Nothing)
		Call ASPUnit.Same(objModule.Lifecycle, Nothing, "CreateModule factory method should return ASPUnitModule with specified lifecycle")
		Set objModule = Nothing
	End Sub

	' Test that ASPUnitTester CreateTest return expected objects

	Sub ASPUnitTesterCreateTest()
		Dim objTest

		Set objTest = objService.CreateTest("Test")
		Call ASPUnit.Equal(TypeName(objTest), "ASPUnitTest", "CreateTest factory method should return ASPUnitTest")
		Set objTest = Nothing
	End Sub

	Sub ASPUnitTesterCreateTestName()
		Dim objTest

		Set objTest = objService.CreateTest("Test")
		Call ASPUnit.Equal(objTest.Name, "Test", "CreateTest factory method should return ASPUnitTest with specified name")
		Set objTest = Nothing
	End Sub

	' Test that ASPUnitTester CreateLifecycle return expected objects

	Sub ASPUnitTesterCreateLifecycle()
		Dim objLifecycle

		Set objLifecycle = objService.CreateLifecycle("", "")
		Call ASPUnit.Equal(TypeName(objLifecycle), "ASPUnitTestLifecycle", "CreateLifecycle factory method should return ASPUnitTestLifecycle")
		Set objLifecycle = Nothing
	End Sub

	Sub ASPUnitTesterCreateLifecycleSetup()
		Dim objLifecycle

		Set objLifecycle = objService.CreateLifecycle("Setup", "")
		Call ASPUnit.Equal(objLifecycle.Setup, "Setup", "CreateTest factory method should return ASPUnitTest with specified setup")
		Set objLifecycle = Nothing
	End Sub

	Sub ASPUnitTesterCreateLifecycleTeardown()
		Dim objLifecycle

		Set objLifecycle = objService.CreateLifecycle("", "Teardown")
		Call ASPUnit.Equal(objLifecycle.Teardown, "Teardown", "CreateTest factory method should return ASPUnitTest with specified teardown")
		Set objLifecycle = Nothing
	End Sub
%>