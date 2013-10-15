<% Option Explicit %>

<!-- #include file="../Lib/ASPUnit.asp" -->

<%
	Dim objLifecycle
	Set objLifecycle = ASPUnit.CreateLifeCycle("Setup", "Teardown")

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitLibrary Run Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitLibraryRunsTestTask"), _
				ASPUnit.CreateTest("ASPUnitLibraryRunsUITask") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.Run()

	' Create a global instance of ASPUnitLibrary for testing

	Sub Setup()
		Call ExecuteGlobal("Dim objLibrary")
		Set objLibrary = New ASPUnitLibrary
	End Sub

	Sub Teardown()
		Set objLibrary = Nothing
	End Sub

	' Create mock services that set global variables to indicate when Run method is executed

	Class ASPUnitLibraryMockTester
		Public Sub Run()
			Call ExecuteGlobal("Dim blnTesterRan")
			blnTesterRan = True
		End Sub
	End Class

	Class ASPUnitLibraryMockRunner
		Public Sub Run()
			Call ExecuteGlobal("Dim blnRunnerRan")
			blnRunnerRan = True
		End Sub
	End Class

	' Test that ASPUnitLibrary implements services on Run

	Sub ASPUnitLibraryRunsTestTask()
		Set objLibrary.Tester = New ASPUnitLibraryMockTester
		objLibrary.Task = "TEST"
		objLibrary.Run()

		Call ASPUnit.Equal(blnTesterRan, True, "Run method should execute test service when TEST task specified")
	End Sub

	Sub ASPUnitLibraryRunsUITask()
		Set objLibrary.Runner = New ASPUnitLibraryMockRunner
		objLibrary.Run()

		Call ASPUnit.Equal(blnRunnerRan, True, "Run method should execute Runner service when task is unspecified")
	End Sub
%>