<% Option Explicit %>

<!-- #include file="../Lib/ASPUnit.asp" -->
<!-- #include file="Include/Account.asp" -->
<!-- #include file="Include/AccountTransferService.asp" -->

<%
	' Register unit test modules (groups of tests)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"AccountTransferService Transfer Tests", _
			Array( _
				ASPUnit.CreateTest("SourceBalanceAfterTransfer"), _
				ASPUnit.CreateTest("DestinationBalanceAfterTransfer") _
			), _
			ASPUnit.CreateLifeCycle("Setup", "Teardown") _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"AccountTransferService Insufficient Funds Tests", _
			Array( _
				ASPUnit.CreateTest("TransferWithInsufficientFundsRaisesError"), _
				ASPUnit.CreateTest("SourceBalanceAfterTransferWithInsufficientFunds"), _
				ASPUnit.CreateTest("DestinationBalanceAfterTransferWithInsufficientFunds") _
			), _
			ASPUnit.CreateLifeCycle("Setup", "Teardown") _
		) _
	)

	Call ASPUnit.Run()

	' Lifecycle methods

	Sub Setup()
		Call ExecuteGlobal("Dim objService")
		Call ExecuteGlobal("Dim objSource")
		Call ExecuteGlobal("Dim objDestination")

		Set objService = New AccountTransferService
		Set objSource = New Account
		Set objDestination = New Account

		objSource.Deposit(100)
		objDestination.Deposit(100)
	End Sub

	Sub Teardown()
		Set objService = Nothing
		Set objSource = Nothing
		Set objDestination = Nothing
	End Sub

	' Test methods

	Function SourceBalanceAfterTransfer()
		Call objService.Transfer(objSource, objDestination, 50)
		Call ASPUnit.Equal(objSource.Balance, 50, "Source balance should decrease by specified transfer amount")
	End Function

	' Intentional fail
	Function DestinationBalanceAfterTransfer()
		Call objService.Transfer(objSource, objDestination, 50)
		Call ASPUnit.Equal(objDestination.Balance, 151, "Destination balance should increase by specified transfer amount")
	End Function

	Function TransferWithInsufficientFundsRaisesError()
		On Error Resume Next
		Call objService.Transfer(objSource, objDestination, 95)
		Call ASPUnit.NotEqual(Err.Number, 0, "Transfer with insufficient funds should raise an error")
		On Error Goto 0
	End Function

	Function SourceBalanceAfterTransferWithInsufficientFunds()
		On Error Resume Next
		Call objService.Transfer(objSource, objDestination, 95)
		Call ASPUnit.Equal(objSource.Balance, 100, "Transfer with insufficient funds should leave source balance unchanged")
		On Error Goto 0
	End Function

	Function DestinationBalanceAfterTransferWithInsufficientFunds()
		On Error Resume Next
		Call objService.Transfer(objSource, objDestination, 95)
		Call ASPUnit.Equal(objDestination.Balance, 100, "Transfer with insufficient funds should leave destination balance unchanged")
		On Error Goto 0
	End Function
%>