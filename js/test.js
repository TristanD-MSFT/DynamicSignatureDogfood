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
			let size = Office.context.roamingSettings.get("signatureFontSize");
			if(size != "")
			{
				$('fontsize').val = size;
			}
			else
			{
				$('fontsize').val = 9;
			}
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

function onNewComposeHandler(eventObj)
{
	//get current time
	let today = new Date();
	let time = today.getHours();
	let day = today.getDay();
  
	if (day == 0 || day == 6 || time < 8 || time > 16) {
	  test_signature();
	}

	eventObj.completed();
}

function test_signature()
{
	Office.context.mailbox.item.body.setSignatureAsync(
		"<p style='margin-bottom:0in;line-height:normal'><span style='font-size:9.0pt'>------------</span></p>" + 
		"<p style='margin-bottom:0in;line-height:normal'><span style='font-size:9.0pt'>Your family and personal time is important to me; after-hours responses not required or expected!</span></p >" +
		"<p></p>",
		{ coercionType: Office.CoercionType.Html }
	);
}

function saveSignatureSize()
{
	let signatureSize = $('fontsize').val;
	Office.context.roamingSettings.set("signatureFontSize",signatureSize.toString());

	// Save settings in the mailbox to make it available in future sessions.
	Office.context.roamingSettings.saveAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
			$('message').val = "Action failed with message: " + result.error.message;
          console.error(`Action failed with message ${result.error.message}`);
        } else {
			$('message').val = "Settings saved with status: " + result.status;
          console.log(`Settings saved with status: ${result.status}`);
        }
      });
}