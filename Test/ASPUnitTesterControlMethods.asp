<% Option Explicit %>

<!-- #include file="../Lib/ASPUnit.asp" -->

<%
	Dim objLifecycle
	Set objLifecycle = ASPUnit.CreateLifeCycle("Setup", "Teardown")

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester AddModule Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterAddModule"), _
				ASPUnit.CreateTest("ASPUnitTesterAddModules") _
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

	' Test that ASPUnitTester adds specified modules

	Sub ASPUnitTesterAddModule()
		Call objService.AddModule(Nothing)
		Call ASPUnit.Equal(objService.Modules.Count, 1, "AddModule method should add module to collection")
	End Sub

	Sub ASPUnitTesterAddModules()
		Call objService.AddModules(Array(Nothing, Nothing))
		Call ASPUnit.Equal(objService.Modules.Count, 2, "AddModules method should add modules to collection")
	End Sub
%>