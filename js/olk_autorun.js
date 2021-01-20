/*
 * @file: olk_autorun.js (used by Win32)
 * @author: Tristan Davis
 */

function onNewComposeHandler(eventObj)
{
  let today = new Date();
  let time = today.getHours();
  let day = today.getDay();

  if (day == 0 || day == 6 || time < 8 || time > 16) {

    let size = Office.context.roamingSettings.get("signatureFontSize");
    console.log(`Loaded roaming setting: ${size}`);
    if (size == "") 
    {
      size = "9";		
    }

    Office.context.mailbox.item.body.setSignatureAsync
    (
      "<p style='margin-bottom:0in;line-height:normal'><span style='font-size:" + size + ".0pt'>------------</span></p>" + 
      "<p style='margin-bottom:0in;line-height:normal'><span style='font-size:" + size + ".0pt'>Your family and personal time is important to me; after-hours responses not required or expected!</span></p >" +
      "<p></p>",
        {
            "coercionType": "html",
            "asyncContext" : eventObj
        },
        function (asyncResult)
        {
            asyncResult.asyncContext.completed({ "key00" : "val00" });
        }
    );
  }
  else {
    eventObj.completed({ "key00" : "val00"});
  }
}
Office.actions.associate("onNewComposeHandler", onNewComposeHandler);