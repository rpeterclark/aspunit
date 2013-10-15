<% Option Explicit %>

<!-- #include file="../Lib/ASPUnit.asp" -->
<!-- #include file="Include/Account.asp" -->

<%
	' Register unit test modules (groups of tests)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"Account Balance Tests", _
			Array( _
				ASPUnit.CreateTest("BalanceAfterDeposit"), _
				ASPUnit.CreateTest("BalanceAfterWithdrawal") _
			), _
			ASPUnit.CreateLifeCycle("Setup", "Teardown") _
		) _
	)

	Call ASPUnit.Run()

	' Lifecycle methods

	Sub Setup()
		Call ExecuteGlobal("Dim objAccount")
		Set objAccount = New Account
	End Sub

	Sub Teardown()
		Set objAccount = Nothing
	End Sub

	' Test methods

	Function BalanceAfterDeposit()
		Call objAccount.Deposit(50)
		Call ASPUnit.Equal(objAccount.Balance, 50, "Balance should increase by specified amount")
	End Function

	Function BalanceAfterWithdrawal()
		Call objAccount.Deposit(100)
		Call objAccount.Withdraw(50)
		Call ASPUnit.Equal(objAccount.Balance, 50, "Balance should decrease by specified amount")
	End Function
%>