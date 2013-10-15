<!-- #include file="classes/ASPUnitLibrary.asp" -->
<!-- #include file="classes/ASPUnitTester.asp" -->
<!-- #include file="classes/ASPUnitRunner.asp" -->
<!-- #include file="classes/ASPUnitUIModern.asp" -->
<!-- #include file="classes/ASPUnitJSONResponder.asp" -->

<%
	Dim ASPUnit
	Set ASPUnit = New ASPUnitLibrary

	ASPUnit.Task = Request.QueryString("task")
%>