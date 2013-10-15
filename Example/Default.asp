<% Option Explicit %>

<!-- Include ASPUnit library -->
<!-- #include file="../Lib/ASPUnit.asp" -->

<%
	' Register pages to test
	Call ASPUnit.AddPage("/Example/TestAccount.asp")
	Call ASPUnit.AddPage("/Example/TestAccountTransferService.asp")

	' Execute tests
	Call ASPUnit.Run()
%>