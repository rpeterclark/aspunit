# ASPUnit

> A Classic ASP Unit Testing Framework

## Why?

Is this framework a decade late? Probably! But there's still work to be done on older platforms. Some places have deep infrastructure built on classic ASP and it isn't getting rewritten anytime soon. May as well use some tools like this so that development on this platform can be a little less painful.

## Getting Started

You have a couple options for getting this framework into your project. First, I suggest creating a `/test` folder within your project to contain the framework and your tests. If you are comfortable with [Git submodules](http://git-scm.com/book/en/Git-Tools-Submodules), you can just clone this repo into your Git repo's test folder. Otherwise you can clone to any location and copy the files to your test folder.

Note that this framework has to run within an IIS web server, like [IIS Express](http://www.microsoft.com/en-us/download/details.aspx?id=1038). If you're using Grunt, you should take a look at [grunt-iisexpress](https://github.com/rpeterclark/grunt-iisexpress).

Check out the `/example` folder within this repo to see a working set of tests. You can also check the `/test` folder where this projects tests itself.

## Usage

Here's a minimal test setup:

```
<!-- #include file="/Test/Lib/ASPUnit.asp" -->

Call ASPUnit.AddModule( _
	ASPUnit.CreateModule( _
		"Simple Tests", _
		Array( _
			ASPUnit.CreateTest("TestSomething"), _
			ASPUnit.CreateTest("TestSomethingElse") _
		), _
		ASPUnit.CreateLifeCycle("Setup", "Teardown") _
	) _
)

Call ASPUnit.Run()

Sub Setup() ' Could do something here
End Sub

Sub Teardown() ' Then undo it here
End Sub

Function TestSomething()
	Call ASPUnit.Ok(True, "This assertion should pass")
End Function

Function TestSomethingElse()
	Call ASPUnit.Ok(False, "This assertion should fail")
End Function
```

## Assertions

- `Ok(value, description)`: Tests if `value` is `True`
- `Equal(actual, expected, description)`: Tests if `actual` is equal to `expected`, uses simple `=` operator.
- `NotEqual(actual, expected, description)`: Tests if `actual` is not equal to `expected`, uses simple `=` operator.
- `Same(actual, expected, description)`: Tests if `actual` refers to the same object as `expected`, uses `Is` operator.
- `NotSame(actual, expected, description)`: Tests if `actual` does not refer to the same object as `expected`, uses `Is` operator.

## Example

Let's say you have a file `/Include/Account.asp` that contains a class to test:

```
<%
Class Account
    Private m_Balance

	Private Sub Class_Initialize()
		m_Balance = 0
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
End Class
%>
```

Add a file `/Test/AccountTests.asp`. Include the ASPUnit framework and the file you want to test:

```
<!-- #include file="/Test/Lib/ASPUnit.asp" -->
<!-- #include file="/Include/Account.asp" -->
```

You might want to write setup and teardown methods for stuff that is common across your tests:

```
Sub Setup()
	Call ExecuteGlobal("Dim objAccount")
	Set objAccount = New Account
End Sub

Sub Teardown()
	Set objAccount = Nothing
End Sub
```

Write some tests for this class:

```
Function BalanceAfterDeposit()
	Call objAccount.Deposit(50)
	Call ASPUnit.Equal(objAccount.Balance, 50, "Balance should increase by specified amount")
End Function

Function BalanceAfterWithdrawal()
	Call objAccount.Deposit(100)
	Call objAccount.Withdraw(50)
	Call ASPUnit.Equal(objAccount.Balance, 50, "Balance should decrease by specified amount")
End Function
```

Now we register these tests with ASPUnit. This is done by registering a `Module` class that contains information about one or many tests. You can register as many modules as you'd like. Modules are a good way of grouping related tests.

```
Call ASPUnit.CreateModule(ModuleName, ArrayOfTests, Lifecycle)
```

So in this case, it would look like this:

```
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
```

Finally, you call the `Run` method:

```
Call ASPUnit.Run()
```

You can now open this page in a browser and see the results of the tests! This is because the page defaults to "runner" mode. Runner mode first writes the UI and then asyncronously calls itself again in "tester" mode, using the JSON results to populate the runner UI. If you are interested in the JSON response, you can add `?task=test` to the page URL.

### Run Lots of Tests

Now have a single page with tests. We don't want that page to have to test *everything* in our project and we don't want to have to open a bunch of different pages to run all of the tests. To accomodate this, you can register URLs with ASPUnit and it will run them all. I would encourage you to create individual files for logical areas that you want to test. So, pages are groups of modules, and modules are groups of tests. It's up to you to determine how your project should use these levels of granularity.

In your test folder, create a new ASP file and write something like this:

```
<!-- #include file="../Lib/ASPUnit.asp" -->

<%
	Call ASPUnit.AddPage("/Test/TestAccount.asp")
	Call ASPUnit.AddPage("/Test/TestAccountTransferService.asp")
	Call ASPUnit.Run()
%>
```

You can then open this page and it will aggregate the test results of each registered page. Neat!

## Included PhantomJS Runner

A PhantomJS-powered headless test runner is included, providing basic console output for ASPUnit tests.

### Usage
```bash
  phantomjs runner.js [url-of-your-aspunit-testsuite]
```

### Example
```bash
  phantomjs runner.js http://localhost:8080/aspunit/test/
```

### Notes
 - Requires [PhantomJS](http://phantomjs.org/) 1.6+ (1.7+ recommended).
 - If you're using Grunt, you should take a look at [grunt-aspunit](https://github.com/rpeterclark/grunt-aspunit).

## Contributing
In lieu of a formal styleguide, take care to maintain the existing coding style. Add unit tests for any new or changed functionality.