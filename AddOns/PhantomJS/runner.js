(function() {
	'use strict';

	var page, url, args = require('system').args;

	if (args.length < 2) {
		console.error('Usage:\n  phantomjs runner.js [url-of-aspunit-tests]');
		phantom.exit(1);
	}

	url = args[1];
	page = require('webpage').create();

	page.onConsoleMessage = function(msg) {
		console.log(msg);
	};

	page.onInitialized = function() {
		page.evaluate(addLogging);
	};

	page.onCallback = function(msg) {
		if (msg.name == 'ASPUnit.Done') {
			phantom.exit(msg.passed ? 0 : 1);
		}
	};

	page.open(url, function(status) {
		if (status != 'success') {
			console.error('Unable to access url: ' + status);
			phantom.exit(1);
		} else {
			var aspunitMissing = page.evaluate(function() { return (typeof ASPUnit === 'undefined' || !ASPUnit); });
			if (aspunitMissing) {
				console.error('The "ASPUnit" object is not present on this page');
				phantom.exit(1);
			}
		}
	});

	function addLogging() {
		window.document.addEventListener('DOMContentLoaded', function() {
			var status = null;
			var pageSuccessCount = 0;

			ASPUnit.onPageSuccess(function(details) {
				pageSuccessCount++;

				if (!details.passed) {
					console.log('\nTest failed: ' + details.page);
					for (var i = 0; i < details.modules.length; i++) {
						for (var j = 0; j < details.modules[i].tests.length; j++) {
							var test = details.modules[i].tests[j];
							if (!test.passed) {
								console.log(test.name + ': ' + test.description);
							}
						}
					}
				}
			});

			ASPUnit.onPageFail(function(details) {
				console.log('\nPage failed: ' + details.page + '\n' + details.error + ': ' + details.description);
			});

			ASPUnit.onPageFinish(function(details) {
				status = details.status;
			});

			ASPUnit.onFinish(function() {
				console.log('\n' + pageSuccessCount + ' of ' + status.pageCount + ' pages tested, ' +
					status.passCount + ' tests passed, ' + (status.testCount - status.passCount) + ' tests failed.');

				window.callPhantom({
					name: 'ASPUnit.Done',
					passed: (status.passCount == status.testCount && status.pageCount == pageSuccessCount)
				});
			});

		}, false);
	}
})();