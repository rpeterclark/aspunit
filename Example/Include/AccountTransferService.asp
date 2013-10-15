<%
	Class AccountTransferService
		Public Sub Transfer(objSource, objDestination, dblAmount)
			If (objSource.Balance - dblAmount) < objSource.MinimumBalance Then
				Call Err.Raise(9999, TypeName(Me), "Insufficient Funds")
			End If

			objSource.Withdraw(dblAmount)
			objDestination.Deposit(dblAmount)
		End Sub
	End Class
%>