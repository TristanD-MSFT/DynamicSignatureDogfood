/*
 * @file: test.js (used by OWA)
 * @author: Tristan Davis
 */
let _output;

Office.initialize = function(reason)
{
	on_initialization_complete();
	console.log("office.js initialized");
}

function on_initialization_complete()
{
	$(document).ready
	(
		function()
		{
			console.log("document ready");
			let size = Office.context.roamingSettings.get("signatureFontSize");
			console.log("Loaded roaming setting: " + size);
			if (size == undefined) {
				console.log("Setting font size to default.");
			  	$("#fontsize").val("9");
			} else {
				console.log("Setting font size to: " + size);
				$("#fontsize").val(size);	
			}

			let startTime = Office.context.roamingSettings.get("startTime");
			console.log("Loaded roaming setting: " + startTime);
			if (startTime == undefined) {
				console.log("Setting start time to default.");
			  	$("#startTime").val("8");
			} else {
				console.log("Setting start time to: " + startTime);
				$("#startTime").val(startTime);	
			}

			let endTime = Office.context.roamingSettings.get("endTime");
			console.log("Loaded roaming setting: " + endTime);
			if (endTime == undefined) {
				console.log("Setting end time to default.");
			  	$("#endTime").val("17");
			} else {
				console.log("Setting end time to: " + endTime);
				$("#endTime").val(endTime);	
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

	let startTime = Office.context.roamingSettings.get("startTime");
	let endTime = Office.context.roamingSettings.get("endTime");
	console.log("Loaded roaming settings");
	if (startTime == "") 
	{
		startTime = 7;		
	}
	if (endTime == "")
	{
		endTime = 17;
	}
  
	if (day == 0 || day == 6 || time < startTime || time >= endTime) {
	  test_signature();
	}

	eventObj.completed();
}

function test_signature()
{
	let size = Office.context.roamingSettings.get("signatureFontSize");
	console.log("Loaded roaming setting: " + size);
	if (size == "") 
	{
		size = "9";		
	}
	insertSignature(size);
}

function insertSignature(fontSize)
{
	Office.context.mailbox.item.body.setSignatureAsync(
		"<p style='margin-bottom:0in;line-height:normal'><span style='font-size:" + fontSize + ".0pt'>------------</span></p>" + 
		"<p style='margin-bottom:0in;line-height:normal'><span style='font-size:" + fontSize + ".0pt'>Your family and personal time is important to me; after-hours responses not required or expected!</span></p >" +
		"<p></p>",
		{ coercionType: Office.CoercionType.Html }
	);
}

function saveSignatureSize()
{
	let startTime = $("#startTime").val();
	let endTime = $("#endTime").val();

	if(parseInt(startTime) >= parseInt(endTime))
	{
		console.log("start time before end time? start time = " + startTime + ", end time = " + endTime);
		$("#errorMessage").val("ERROR: Start time cannot be before end time.");
		return;
	}
	else
	{
		$("#errorMessage").val("");
	}

	let signatureSize = $("#fontsize").val();

	Office.context.roamingSettings.set("signatureFontSize", signatureSize.toString());
	Office.context.roamingSettings.set("startTime", startTime.toString());
	Office.context.roamingSettings.set("endTime", endTime.toString());

	// Save settings in the mailbox to make it available in future sessions.
	Office.context.roamingSettings.saveAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Action failed with message: " + result.error.message);
        } else {
          console.log("Settings saved with status: " + result.status);
        }
      });
}