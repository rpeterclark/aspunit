<% Option Explicit %>

<!-- Include ASPUnit library -->
<!-- #include file="../Lib/ASPUnit.asp" -->

<%
	' Register pages to test
	Call ASPUnit.AddPage("TestAccount.asp")
	Call ASPUnit.AddPage("TestAccountTransferService.asp")

	' Execute tests
	Call ASPUnit.Run()
%>