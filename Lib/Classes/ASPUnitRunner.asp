<%
	Class ASPUnitRunner
		Private _
			m_Theme, _
			m_Pages

		Private Sub Class_Initialize()
			Set m_Theme = New ASPUnitUIModern
			Set m_Pages = Server.CreateObject("System.Collections.ArrayList")
		End Sub

		Private Sub Class_Terminate()
			Set m_Pages = Nothing
			Set m_Theme = Nothing
		End Sub

		Public Property Set Theme(ByRef objValue)
			Set m_Theme = objValue
		End Property

		Public Property Get Pages
			Set Pages = m_Pages
		End Property

		' Public methods to specify test pages

		Public Sub AddPage(strPage)
			Call m_Pages.Add(strPage)
		End Sub

		Public Sub AddPages(arrPages)
			Dim i

			For i = 0 To UBound(arrPages)
				Call AddPage(arrPages(i))
			Next
		End Sub

		' Method to run UI

		Public Sub Run()
			If m_Pages.Count = 0 Then
				Call AddCurrentPage()
			End If

			Call m_Theme.Render(Me)
		End Sub

		Private Sub AddCurrentPage()
			Call AddPage(Request.ServerVariables("URL"))
		End Sub

		Public Sub RenderJSLib() %>
			<script>
				var ASPUnit = function() {
					'use strict';

					var config = {
						pages: [],
						callbacks: {
							onStart: [],
							onStop: [],
							onPageStart: [],
							onPageSuccess: [],
							onPageFail: [],
							onPageFinish: [],
							onFinish: [],
						}
					}

					var getPageTimeout = null,
						getPageXHR = null;

					var status = {
						pageIndex: 0,
						pageCount: 0,
						testCount: 0,
						passCount: 0
					};

					function registerCallback(key, callback) {
						config.callbacks[key].push(callback);
					}

					function callback(key, args) {
						var callbacks = config.callbacks[key];
						for (var i = 0, len = callbacks.length; i < len; i++) {
							callbacks[i].call({}, args);
						}
					}

					function getPage(page) {
						callback('onPageStart', {page: page});

						getPageXHR = $.getJSON(page + '?task=test')
						.done(function(data) {
							status.pageDoneCount++;
							status.testCount += data.testCount;
							status.passCount += data.passCount;

							callback('onPageSuccess', $.extend({page: page}, data));
						})
						.fail(function(jqXHR, textStatus, errorThrown) {
							callback('onPageFail', {
								page: page,
								error: errorThrown,
								description: (jqXHR.status != 404) ? jqXHR.responseText : ''
							});
						})
						.always(function() {
							callback('onPageFinish', {
								page: page,
								status: status
							});

							if (status.pageIndex < (config.pages.length - 1)) {
								getPageTimeout = setTimeout(function() {
									getPage(config.pages[++status.pageIndex]);
								}, 100);
							} else {
								callback('onFinish');
							}
						});
					}

					return {
						load: function(pages) {
							config.pages = pages;
							status.pageCount = config.pages.length;
							this.start();
						},

						start: function() {
							callback('onStart');
							getPage(config.pages[status.pageIndex]);
						},

						stop: function() {
							getPageXHR.abort();
							clearTimeout(getPageTimeout);
							callback('onStop');
						},

						onStart: function(callback) { registerCallback('onStart', callback); },
						onPageStart: function(callback) { registerCallback('onPageStart', callback); },
						onPageSuccess: function(callback) { registerCallback('onPageSuccess', callback); },
						onPageFail: function(callback) { registerCallback('onPageFail', callback); },
						onPageFinish: function(callback) { registerCallback('onPageFinish', callback); },
						onFinish: function(callback) { registerCallback('onFinish', callback); }
					};
				}();
			</script> <%
		End Sub

		Public Sub RenderJSInit() %>
			<script>$(function(){ASPUnit.load([<%= GetPagesAsJSString() %>])});</script> <%
		End Sub

		Private Function GetPagesAsJSString()
			Dim strReturn, _
				i

			strReturn = ""
			For i = 0 To (m_Pages.Count - 1)
				strReturn = strReturn & "'" & m_Pages.Item(i) & "'"
				If i < (m_Pages.Count - 1) Then
					strReturn = strReturn & ", "
				End If
			Next

			GetPagesAsJSString = strReturn
		End Function
	End Class
%>