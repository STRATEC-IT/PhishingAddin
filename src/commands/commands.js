/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global global, Office, self, window */
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

function sucessNotif(msg) {
  var id = "0";
  var details = {
    type: "informationalMessage",
    icon: "icon16",
    message: msg,
    persistent: false
  };
  Office.context.mailbox.item.notificationMessages.addAsync(id, details, function(value) {});
}

function failedNotif(msg) {
  var id = "0";
  var details = {
    type: "informationalMessage",
    icon: "icon16",
    message: msg,
    persistent: false
  };
  Office.context.mailbox.item.notificationMessages.addAsync(id, details, function(value) {});
}

function getItemRestId() {
  if (Office.context.mailbox.diagnostics.hostName === "OutlookIOS") {
    // itemId is already REST-formatted.
    return Office.context.mailbox.item.itemId;
  } else {
    // Convert to an item ID for API v2.0.
    return Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  }
}

// Variables that we'll use to communicate with EWS
var item_id;
var mailbox;

// This function handles the click event of the sendNow button.
// It retrieves the current mail item, so that we can get its itemId property.
// It also retrieves the mailbox, so that we can make an EWS request
// to get more properties of the item. In our case, we are interested in the ChangeKey
// property, becuase we need that to forward a mail item.
function simpleForwardEmail() {
	var item = Office.context.mailbox.item;
	item_id = item.itemId;
	mailbox = Office.context.mailbox;
	sucessNotif("start simpleForwardEmail");
	// The following string is a valid SOAP envelope and request for getting the properties
	// of a mail item. Note that we use the item_id value (which we obtained above) to specify the item
	// we are interested in.
	var soapToGetItemData = '<?xml version="1.0" encoding="utf-8"?>' +
		'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
		'               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
		'               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
		'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
		'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
		'  <soap:Header>' +
		'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
		'  </soap:Header>' +
		'  <soap:Body>' +
		'    <GetItem' +
		'                xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
		'                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
		'      <ItemShape>' +
		'        <t:BaseShape>IdOnly</t:BaseShape>' +
		'      </ItemShape>' +
		'      <ItemIds>' +
		'        <t:ItemId Id="' + item_id + '"/>' +
		'      </ItemIds>' +
		'    </GetItem>' +
		'  </soap:Body>' +
		'</soap:Envelope>';
	sucessNotif("end simpleForwardEmail");
	// The makeEwsRequestAsync method accepts a string of SOAP and a callback function
	mailbox.makeEwsRequestAsync(soapToGetItemData, soapToGetItemDataCallback);
}

// This function is the callback for the makeEwsRequestAsync method
// In brief, it first checks for an error repsonse, but if all is OK
// it then parses the XML repsonse to extract the ChangeKey attribute of the 
// t:ItemId element.
function soapToGetItemDataCallback(asyncResult) {
	var parser;
	var xmlDoc;

	if (asyncResult.error != null) {
		app.showNotification("EWS Status", asyncResult.error.message);            
	}
	else {
		var response = asyncResult.value;
		if (window.DOMParser) {
			var parser = new DOMParser();
			xmlDoc = parser.parseFromString(response, "text/xml");
		}
		else // Older Versions of Internet Explorer
		{
			xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
			xmlDoc.async = false;
			xmlDoc.loadXML(response);
		}
		var changeKey = xmlDoc.getElementsByTagName("t:ItemId")[0].getAttribute("ChangeKey");

		// Now that we have a ChangeKey value, we can use EWS to forward the mail item.
		// The first thing we'll do is get an array of email addresses that the user
		// has typed into the To: text box.
		// We'll also get the comment that the user may have provided in the Comment: text box.
		// var toAddresses = document.getElementById("groupEmails").value;
		// var addresses = toAddresses.split(";");
		var addressesSoap = "<t:Mailbox><t:EmailAddress>m.schlehuber@stratec.com</t:EmailAddress></t:Mailbox>";

		// The following loop build an XML fragment that we will insert into the SOAP message
		//for (var address = 0; address < addresses.length; address++) {
		//	addressesSoap += "<t:Mailbox><t:EmailAddress>" + addresses[address] + "</t:EmailAddress></t:Mailbox>";
		//}
		//var comment = document.getElementById("groupComment").value;
		var comment = "Phishing Test";
		
		// The following string is a valid SOAP envelope and request for forwarding
		// a mail item. Note that we use the item_id value (which we obtained in the click event handler)
		// to specify the item we are interested in,
		// along with its ChangeKey value that we have just determined near the top of this function.
		// We also provide the XML fragment that we built in the loop above to specify the recipient addresses,
		// and the comment that the user may have provided in the Comment: text box
		var soapToForwardItem = '<?xml version="1.0" encoding="utf-8"?>' +
			'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
			'               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
			'               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
			'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
			'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
			'  <soap:Header>' +
			'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
			'  </soap:Header>' +
			'  <soap:Body>' +
			'    <m:CreateItem MessageDisposition="SendAndSaveCopy">' +
			'      <m:Items>' +
			'        <t:ForwardItem>' +
			'          <t:ToRecipients>' + addressesSoap + '</t:ToRecipients>' +
			'          <t:ReferenceItemId Id="' + item_id + '" ChangeKey="' + changeKey + '" />' +
			'          <t:NewBodyContent BodyType="Text">' + comment + '</t:NewBodyContent>' +
			'        </t:ForwardItem>' +
			'      </m:Items>' +
			'    </m:CreateItem>' +
			'  </soap:Body>' +
			'</soap:Envelope>';

		// As before, the makeEwsRequestAsync method accepts a string of SOAP and a callback function.
		// The only difference this time is that the body of the SOAP message requests that the item
		// be forwarded (rather than retrieved as in the previous method call)
		mailbox.makeEwsRequestAsync(soapToForwardItem, soapToForwardItemCallback);
	}
}

// This function is the callback for the above makeEwsRequestAsync method
// In brief, it first checks for an error repsonse, but if all is OK
// it then parses the XML repsonse to extract the m:ResponseCode value.
function soapToForwardItemCallback(asyncResult) {
	var parser;
	var xmlDoc;

	if (asyncResult.error != null) {
		app.showNotification("EWS Status", asyncResult.error.message);
	}
	else {
		var response = asyncResult.value;
		if (window.DOMParser) {
			parser = new DOMParser();
			xmlDoc = parser.parseFromString(response, "text/xml");
		}
		else // Older Versions of Internet Explorer
		{
			xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
			xmlDoc.async = false;
			xmlDoc.loadXML(response);
		}

		// Get the required response, and if it's NoError then all has succeeded, so tell the user.
		// Otherwise, tell them what the problem was. (E.G. Recipient email addresses might have been
		// entered incorrectly --- try it and see for yourself what happens!!)
		var result = xmlDoc.getElementsByTagName("m:ResponseCode")[0].textContent;
		if (result == "NoError") {
			sucessNotif("Success!");
		}
		else {
			failedNotif("The following error code was recieved: " + result);
		}
	}
}

/* Simple Forward */
function simpleForwardEmail_Rest() {
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function(result) {
    var ewsId = Office.context.mailbox.item.itemId;
    var accessToken = result.value;
    simpleForwardFunc(accessToken);
  });
  
}

function simpleForwardFunc(accessToken) {
  var itemId = getItemRestId();

  // Construct the REST URL to the current item.
  // Details for formatting the URL can be found at
  // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
  var forwardUrl = Office.context.mailbox.restUrl + "/v1.0/me/messages/" + itemId + "/forward";

  const forwardMeta = JSON.stringify({
    Comment: "FYI",
    ToRecipients: [
      {
        EmailAddress: {
          Name: "Test",
          Address: "m.schlehuber@stratec.com"
        }
      }
    ]
  });

  $.ajax({
    url: forwardUrl,
    type: "POST",
    dataType: "json",
    contentType: "application/json",
    data: forwardMeta,
    headers: { Authorization: "Bearer " + accessToken }
  }).always(function(response){
    sucessNotif("Email Forward successful!");
  });
}


/* Forward as Attachment */
function forwardAsAttachment(){
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function(result) {
    var ewsId = Office.context.mailbox.item.itemId;
    var accessToken = result.value;
    forwardAsAttachmentFunc(accessToken);
  });
}

function forwardAsAttachmentFunc(accessToken) {
  var itemId = getItemRestId();
  var getAnItemUrl = Office.context.mailbox.restUrl + "/v1.0/me/messages/" + itemId;
  var sendItemUrl = Office.context.mailbox.restUrl + "/v1.0/me/sendmail";

  $.ajax({
    url: getAnItemUrl,
    type: "GET",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + accessToken }
  }).done(function (responseItem) {
    // #microsoft.graph.message
    // microsoft.graph.outlookItem
    responseItem['@odata.type'] = "#microsoft.graph.message";
    
    /* Now send mail */
    const sendMeta = JSON.stringify({
      "Message": {
        "Subject": "Please Check for Phish Activities!",
        "Body": {
          "ContentType": "Text",
          "Content": "Please Check for Phish Activities and let us know!"
        },
        "ToRecipients": [{
          "EmailAddress": {
            "Address": "m.schlehuber@stratec.com"
          }
        }],
        "Attachments": [
          {
            "@odata.type": "#Microsoft.OutlookServices.ItemAttachment",
            // #Microsoft.OutlookServices.ItemAttachment - worked with graph explorer
            // #Microsoft.graph.ItemAttachment - from stack overfloow
            "Name": responseItem.Subject,
            "Item": responseItem
          }
        ]
      },
      "SaveToSentItems": "false"
    }); // Json.stringify ends

    $.ajax({
      url: sendItemUrl,
      type: "POST",
      dataType: "json",
      contentType: "application/json",
      data: sendMeta,
      headers: { Authorization: "Bearer " + accessToken }
    }).done(function (response) {
      sucessNotif("Email forward as attachment successful!");
    }).fail(function(response){
      failedNotif(response);
    }); // ajax of send mail ends

  }); // ajax.get.done ends
}
