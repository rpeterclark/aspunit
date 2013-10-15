<%
	Class ASPUnitUIModern
		Public Sub Render(ByRef objRunner) %>
			<html>
				<head>
					<title>ASPUnit</title>

					<link href="//cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.0.0/css/bootstrap.min.css" rel="stylesheet">
					<link href="//cdnjs.cloudflare.com/ajax/libs/font-awesome/3.2.1/css/font-awesome.min.css" rel="stylesheet">
					<link href='//fonts.googleapis.com/css?family=Open+Sans:400,600,700' rel='stylesheet' type='text/css'>

					<style type="text/css">
						body {
							background-color: #ECF0F1;
							padding-top: 280px;
							font-family: 'Open Sans', sans-serif;
							-webkit-font-smoothing: antialiased;
						}

						h1, h2, h3, h4, h5, h6 {
							font-family: 'Open Sans', sans-serif;
						}

						a {
							color: #95a5a6;
							transition: all 0.25s;
						}

						a:hover {
							color: #fff;
							text-decoration: none;
						}

						.project-name {
							color: #fff !important;
						}

						#footer .project-name {
							margin-top: 0;
							margin-bottom: 0;
						}

						#container {
							position: relative;
							height: auto;
							min-height: 100%;
						}

						#header {
							color: #ecf0f1;
							background-color: #13202c;
							transition: all 0.5s;
							position: fixed;
							left: 0;
							right: 0;
							top: 0;
							z-index: 1000;
							text-shadow: 0px 2px 1px #000;
						}

						#header a {
							color: #ecf0f1;
						}

						#header a:active {
							position: relative;
							top: 1px;
							text-shadow: 0px 1px 1px #000;
						}

						#header.affix {
							padding-top: 0;
							padding-bottom: 0.2em;
						}

						#main {
							padding-bottom: 12em;
						}

						#footer {
							position: absolute;
							bottom: 0;
							width: 100%;
							padding-top: 2em;
							padding-bottom: 2em;
							margin-top: 4em;
							color: #95a5a6;
							background-color: #13202c;
						}

						.progress-loading {
							box-shadow: 0 2px 6px RGBA(0,0,0,0.25);
							border-bottom: 2px #444 solid;
							border-radius: 0 0 2px 2px;
						}

						.pages-overview {
							position: relative;
						}

						.pages-status {
							display: inline-block;
						}

						.pages-options {
							display: inline-block;
							position: absolute;
							right: 0;
						}

						.page-report .page-overview {
							color: #7F8C8D;
						}

						.page-report .page-overview small {
							color: #95A5A6;
						}

						.page-module {
							margin-bottom: 1em;
							box-shadow: 0 2px 6px RGBA(0,0,0,0.25);
						}

						.page-module .page-module-header {
							color: #fff;
							padding: 1em 2em;
						}

						.page-module-pass .page-module-header {
							background-color: #2ECC71;
							border: 1px #27AE60 solid;
						}

						.page-module-fail .page-module-header {
							background-color: #E74C3C;
							border: 1px #C0392B solid;
						}

						.page-module-name {
							font-weight: bold;
							margin: 0;
						}

						.page-module-tests {
							border-left: 1px #95A5A6 solid;
							border-right: 1px #95A5A6 solid;
							border-bottom: 3px #95A5A6 solid;
							border-radius: 0 0 3px 3px;
						}

						.page-module-test {
							padding: 0.8em 2em;
							background-color: #fff;
							margin-bottom: 1px;
							position: relative;
						}

						.page-module-test-icon {
							position: absolute;
							top: 0.5em;
							left: -11px;
							color: #fff;
							width: 22px;
							height: 22px;
							padding: 3px 4px;
							border-radius: 11px;
						}

						.page-module-test-icon-pass {
							box-shadow: 0 3px 0 #27AE60, 0 6px 3px RGBA(0,0,0,0.25);
							background-color: #2ECC71;
						}

						.page-module-test-icon-fail {
							box-shadow: 0 3px 0 #C0392B, 0 6px 3px RGBA(0,0,0,0.25);
							background-color: #E74C3C;
						}

						.page-module-test-name {
							font-weight: bold;
						}

						.page-module-test-fail .page-module-test-name, .page-module-test-fail .page-module-test-description {
							color: #C0392B;
						}

						@media (max-width: 768px) {
							body {
								padding-top: 260px;
							}

							#main {
								padding-bottom: 15em;
							}

							.pages-options {
								display: inline-block;
								position: relative;
							}
						}
					</style>
				</head>
				<body>
					<div id="container">
						<div id="header" class="jumbotron" data-spy="affix" data-offset-top="10">
							<div class="container">
								<h1 class="project-name"><strong>ASP</strong>Unit</h1>

								<div class="pages-overview">
									<div class="pages-status"></div>

									<div class="pages-options">
										<div class="pages-option-passed-tests">
											<a href="#" class="action-hide-passed"><i class="glyphicon glyphicon-remove-sign"></i> Hide Passed Tests</a>
										</div>
									</div>

									<div class="progress progress-striped progress-loading">
										<div class="progress-bar progress-bar-success" role="progressbar"></div>
										<div class="progress-bar progress-bar-danger" role="progressbar"></div>
									</div>

									<div class="alert alert-danger progress-error" style="display: none;"></div>
								</div>
							</div>
						</div>

						<div id="main">
							<div class="page-reports"></div>
						</div>

						<div id="footer">
							<div class="container">
								<div class="row">
									<div class="col-md-12">
										<h3 class="project-name"><strong>ASP</strong>Unit</h3>
									</div>
								</div>
								<div class="row">
									<div class="col-md-4">
										Classic ASP Unit Testing Library<br />
									</div>
									<div class="col-md-4">
										<ul class="list-unstyled">
											<li><a href="https://github.com/rpeterclark/aspunit/"><i class="icon-github"></i> GitHub Project</a></li>
											<li><a href="https://github.com/rpeterclark/aspunit/wiki/"><i class="icon-book"></i> Documentation</a></li>
											<li><a href="https://github.com/rpeterclark/aspunit/issues/"><i class="icon-bug"></i> Issues</a></li>
										</ul>
									</div>
									<div class="col-md-4">
										MIT Licensed
									</div>
								</div>
							</div>
						</div>
					</div>

					<script id="page-status-template" type="text/x-handlebars-template">
						{{pageNumber}} of {{pageCount}} pages tested, {{passCount}} of {{testCount}} tests passed
					</script>

					<script id="page-report-template" type="text/x-handlebars-template">
						<div class="page-report container">
							<div class="page-overview">
								<h3>{{page}} <small>{{passCount}} of {{testCount}} tests passed</small></h3>
							</div>
							{{#each modules}}
								<div class="page-module page-module-{{#if passed}}pass{{else}}fail{{/if}}">
									<div class="page-module-header">
										<h4 class="page-module-name">{{name}}</h4>
										{{passCount}} of {{testCount}} tests passed, {{failCount}} failed, completed in {{duration}} milliseconds
									</div>

									<div class="page-module-tests">
										{{#each tests}}
											<div class="page-module-test page-module-test-{{#if passed}}pass{{else}}fail{{/if}}">
												{{#if passed}}
													<i class="page-module-test-icon page-module-test-icon-pass glyphicon glyphicon-ok"></i>
												{{else}}
													<i class="page-module-test-icon page-module-test-icon-fail glyphicon glyphicon-remove"></i>
												{{/if}}
												<span class="page-module-test-name">{{name}}:</span>
												<span class="page-module-test-description">{{description}}</span>
											</div>
										{{/each}}
									</div>
								</div>
							{{/each}}
						</div>
					</script>

					<script id="page-error-template" type="text/x-handlebars-template">
						<div class="page-report container">
							<div class="page-module page-module-fail">
									<div class="page-module-header">
										<h4 class="page-module-name">{{page}}</h4>
									</div>

									<div class="page-module-tests">
										<div class="page-module-test page-module-test-fail">
											<i class="page-module-test-icon page-module-test-icon-fail glyphicon glyphicon-remove"></i>
											<span class="page-module-test-name">{{error}}{{#if description}}:{{/if}}</span>
											<span class="page-module-test-description">{{description}}</span>
										</div>
									</div>
								</div>
							</div>
						</div>
					</script>

					<script src="//cdnjs.cloudflare.com/ajax/libs/jquery/2.0.3/jquery.min.js"></script>
					<script src="//cdnjs.cloudflare.com/ajax/libs/jqueryui/1.10.3/jquery-ui.min.js"></script>
					<script src="//cdnjs.cloudflare.com/ajax/libs/handlebars.js/1.0.0/handlebars.min.js"></script>
					<script src="//cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.0.0/js/bootstrap.min.js"></script>

					<%= objRunner.RenderJSLib() %>

					<script>
						$(function() {
							var successCount = 0;
							var pageReportsContainer = $('.page-reports');
							var pageReportTemplate = Handlebars.compile($('#page-report-template').html());
							var pageErrorTemplate = Handlebars.compile($('#page-error-template').html());
							var pageStatusContainer = $('.pages-status');
							var pageStatusTemplate = Handlebars.compile($('#page-status-template').html());
							var pageReportsProgress = $('.progress-loading');
							var progressError = $('.progress-error');

							ASPUnit.onStart(function() {
								pageReportsProgress.addClass('active');
							});

							ASPUnit.onPageSuccess(function(details) {
								successCount++;
								$(pageReportTemplate(details)).appendTo(pageReportsContainer).hide().fadeIn();
							});

							ASPUnit.onPageFail(function(details) {
								$(pageErrorTemplate(details)).appendTo(pageReportsContainer);
							});

							ASPUnit.onPageFinish(function(details) {
								var status = $.extend(details.status, {pageNumber: (details.status.pageIndex + 1)});

								pageStatusContainer.html(pageStatusTemplate(status));

								var progressPct = (status.pageNumber / status.pageCount);
								var passPct = (status.passCount / status.testCount) * (successCount / status.pageNumber);
								var failPct = (1 - passPct);

								pageReportsProgress.find('.progress-bar-success').css({"width": ((passPct * 100) * progressPct) + "%"})
								pageReportsProgress.find('.progress-bar-danger').css({"width": ((failPct * 100) * progressPct) + "%"})
							});

							ASPUnit.onFinish(function() {
								pageReportsProgress.removeClass('active');
							});

							$(document).on('click', '.action-hide-passed', function(e) {
								e.preventDefault();

								var $el = $(e.target);

								if ($el.hasClass('active')) {
									$el.removeClass('active').html('<i class="glyphicon glyphicon-remove-sign"></i> Hide Passed Tests</a>');
									showPassedTests();
								} else {
									$el.addClass('active').html('<i class="glyphicon glyphicon-ok-sign"></i> Show Passed Tests</a>');
									hidePassedTests();
								}
							});

							function hidePassedTests() {
								$('.page-report').each(function() {
									if ($(this).find('.page-module-test-fail').length > 0) {
										$(this).find('.page-module-pass').addClass('hidden');
										$(this).find('.page-module-test-pass').addClass('hidden');
									} else {
										$(this).addClass('hidden');
									}
								});
							}

							function showPassedTests() {
								$('.hidden').removeClass('hidden');
							}
						});
					</script>

					<%= objRunner.RenderJSInit() %>
				</body>
			</html> <%
		End Sub
	End Class
%>