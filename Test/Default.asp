<% Option Explicit %>

<!-- #include file="../Lib/ASPUnit.asp" -->

<%
	' Register pages to test
	Call ASPUnit.AddPages(Array( _
		"/test/ASPUnitLibrary.asp", _
		"/test/ASPUnitTesterFactoryMethods.asp", _
		"/test/ASPUnitTesterAssertionMethods.asp", _
		"/test/ASPUnitTesterControlMethods.asp", _
		"/test/ASPUnitTesterBehaviors.asp", _
		"/test/ASPUnitRunner.asp" _
	))

	' Execute tests
	Call ASPUnit.Run()
%>