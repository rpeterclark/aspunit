<%
	Class Account
		Private _
			m_Balance, _
			m_MinimumBalance

		Private Sub Class_Initialize()
			m_Balance = 0
			m_MinimumBalance = 10
		End Sub

		Public Sub Deposit(dblAmount)
			m_Balance = m_Balance + dblAmount
		End Sub

		Public Sub Withdraw(dblAmount)
			m_Balance = m_Balance - dblAmount
		End Sub

		Public Property Get Balance
			Balance = m_Balance
		End Property

		Public Property Get MinimumBalance
			MinimumBalance = m_MinimumBalance
		End Property
	End Class
%>