/*
 * @file: test.js
 * @author: Microsoft
 */
let _output;

Office.initialize = function(reason)
{
	on_initialization_complete();
}

function on_initialization_complete()
{
	$(document).ready
	(
		function()
		{
			$("button").each
			(
				function()
				{
					$(this).removeAttr("disabled");
				}
			);

			$("textarea").each
			(
				function()
				{
					$(this).removeAttr("disabled");
				}
			);

			_output = $("textarea#output");

			$("div#api_status").html("API Ready!");
		}
	);
}

function showMessage(message)
{
    _output.val(JSON.stringify(message));
}

function show_message_test()
{
	showMessage("hello there! test successful!");
}

// Don't rename this function "clear()" because that is "document.clear"
function clear_output()
{
	_output.val("");
}

function eval_output()
{
	eval(_output.val());
}

function test_signature()
{
	let signature_str = (_output.val()).trim();

	if (signature_str !== "")
	{
		Office.context.mailbox.item.body.setSignatureAsync
		(
			signature_str,

			{
				"coercionType": "html"
			},

			function (asyncResult)
			{
				showMessage(asyncResult);
			}
		);
	}
	else
	{
		showMessage("Enter signature string here!");
	}
}

function onNewComposeHandler(eventObj)
{
	let d = new Date();

	Office.context.mailbox.item.body.setSignatureAsync
	(
		"This is an awesome signature! at " + d.toString(),
		{
			"coercionType": "html",
			"asyncContext" : eventObj
		},
		function (asyncResult)
		{
			asyncResult.asyncContext.completed();
		}
	);
}