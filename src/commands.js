/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global global, Office, self, window */

var PhshingTrainingDomainsSophos = ["bankfraudalerts.com", "buildingmgmt.info", "corporate-realty.co", "court-notices.com", "e-billinvoices.com", "e-documentsign.com", "e-faxsent.com", "e-receipts.co", "epromodeals.com", "fakebookalerts.live", "global-hr-staff.com", "gmailmsg.com", "helpdesk-tech.com", "hr-benefits.site", "it-supportdesk.com", "linkedn.co", "mail-sender.online", "myhr-portal.site", "online-statements.site", "outlook-mailer.com", "secure-bank-alerts.com", "shipping-updates.com", "tax-official.com", "toll-citations.com", "trackshipping.online", "voicemailbox.online", "parking-services.hr-benefits.site" ];

var PhshingTrainingDomainsKaspersky = ["accommodationstravel.com", "avviso-archiviazione.it", "bestjobs.solutions", "blockchain-info.live", "business-information.me", "business-information.store", "corp-email.info", "correo-interno.es", "courrier-interne.fr", "delivery-post.me", "docs-edit.online", "ecalendar.ws", "e-calendario.es", "events-calendar.site", "events-calendar.today", "free-clinics.co", "google-calendar.com", "install-soft.me", "internal-mail.com", "interne-mail.de", "justmailweb.com", "kaspersky.today", "kasperskygroup.com", "kreditbezahlen.de", "lkea.online", "marketingservice.today", "medcenter.world", "medical-help.social", "mydeliverypost.com", "official-inbox.com", "official-law.site", "parties.agency", "paybill.email", "posta-interna.it", "postelivraison.fr", "postoffice.one", "share-to.me", "shop-delivery.store", "soft-exchange.com", "state-official.info", "stop-covid.center", "storagealert.work", "taxpay365.com", "thedeliverypost.com", "top-programme.de", "vosmarchandises.fr"];

var STRATEC_IT_WhiteList = ["safenetid.com", "onlyfy.io", "staysafe.sophos.com"];

// var STRATEC_IT_WhiteList_MailAddresses = ["helpdesk@stratec.com"];

// Display Status messages to the user
var app = (function () {
    "use strict";

    var app = {};

    // Common initialization function (to be called from each page)
    app.initialize = function () {
        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding ms-font-m">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header"></div>' +
                    '<div id="notification-message-body"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });


        // After initialization, expose a common notification function
        app.showNotification = function (header, text) {
            $('#notification-message-header').text(header);
            $('#notification-message-body').text(text);
            $('#notification-message').slideDown('fast');
        };
    };

    return app;
})();



function changeLanguage() {
	if (document.getElementById("content-ger").classList.contains("hidden")) {
		document.getElementById("content-ger").classList.remove("hidden");
		document.getElementById("content-eng").classList.add("hidden");
	} else {
		document.getElementById("content-ger").classList.add("hidden");
		document.getElementById("content-eng").classList.remove("hidden");
	}
}

function trainingDetected() {
	document.getElementById("content-main-eng").classList.add("hidden");
	document.getElementById("content-main-ger").classList.add("hidden");
	document.getElementById("content-training-eng").classList.remove("hidden");
	document.getElementById("content-training-ger").classList.remove("hidden");
}

function legitimateDetected() {
	document.getElementById("content-main-eng").classList.add("hidden");
	document.getElementById("content-main-ger").classList.add("hidden");
	document.getElementById("content-legitimate-eng").classList.remove("hidden");
	document.getElementById("content-legitimate-ger").classList.remove("hidden");
}

function mailSuccessfullyReported() {
	document.getElementById("content-main-eng").classList.add("hidden");
	document.getElementById("content-main-ger").classList.add("hidden");
	document.getElementById("content-reporting-eng").classList.remove("hidden");
	document.getElementById("content-reporting-ger").classList.remove("hidden");
}

function escapeHTML(htmlStr) {
   return htmlStr.replace(/&/g, "&amp;")
         .replace(/</g, "&lt;")
         .replace(/>/g, "&gt;")
         .replace(/"/g, "&quot;")
         .replace(/'/g, "&#39;");        

}




(function () {
    //"use strict";
	
    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $('#report-eng').click(function () { sendNow(); });
			$('#report-ger').click(function () { sendNow(); });
			$('#changeLanguage-DE').click(function () { changeLanguage(); });
			$('#changeLanguage-EN').click(function () { changeLanguage(); });
			$('#delete-eng1').click(function () { deleteMail(); });
			$('#delete-eng2').click(function () { deleteMail(); });
			$('#delete-ger1').click(function () { deleteMail(); });
			$('#delete-ger2').click(function () { deleteMail(); });
			
        });
    };
  
    // This function handles the click event of the sendNow button.
    // It retrieves the current mail item, so that we can get its itemId property.
    // It also retrieves the mailbox, so that we can make an EWS request
    // to get more properties of the item. In our case, we are interested in the ChangeKey
    // property, becuase we need that to forward a mail item.
    function sendNow() {
        var item = Office.context.mailbox.item;
        item_id = item.itemId;
        mailbox = Office.context.mailbox;

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
            '        <t:BaseShape>AllProperties</t:BaseShape>' +
            '      </ItemShape>' +
            '      <ItemIds>' +
            '        <t:ItemId Id="' + item_id + '"/>' +
            '      </ItemIds>' +
            '    </GetItem>' +
            '  </soap:Body>' +
            '</soap:Envelope>';

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
            app.showNotification("Status", asyncResult.error.message);            
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
				// xmlDoc = new ActiveXObject("Msxml2.DOMDocument.3.0");
                xmlDoc.async = false;
                xmlDoc.loadXML(response);
            }
			
            var changeKey = xmlDoc.getElementsByTagName("t:ItemId")[0].getAttribute("ChangeKey");
			var fromAddress = xmlDoc.getElementsByTagName("t:From")[0].getElementsByTagName("t:EmailAddress")[0].textContent;
			
			// get message header for further research
			var header = xmlDoc.getElementsByTagName("t:InternetMessageHeaders")[0];
			var headerText = "";
			for (let i = 0; i < header.childElementCount; i++) {
				headerText = headerText + header.childNodes[i].getAttribute("HeaderName") + ": " + header.childNodes[i].textContent + "\n";
			}
			
			//var test1 = xmlDoc.getElementsByTagName("t:From")[0].getElementsByTagName("t:EmailAddress")[0].textContent;
			//var test3 = test1[0];
			//var test2 = test1.getElementsByTagName("t:EmailAddress")[0];
			//tesasdf;
			//app.showNotification("Status", fromAddress.split('@')[1].trim().toUpperCase());
			// check if sender domain matches PhshingTrainingDomains
			for (let i = 0; i < PhshingTrainingDomainsSophos.length; i++) {
				if (fromAddress.split('@')[1].trim().toUpperCase() === PhshingTrainingDomainsSophos[i].trim().toUpperCase()) {
					trainingDetected();
					return;
				}
			}
			
			for (let i = 0; i < PhshingTrainingDomainsKaspersky.length; i++) {
				if (fromAddress.split('@')[1].trim().toUpperCase() === PhshingTrainingDomainsKaspersky[i].trim().toUpperCase()) {
					trainingDetected();
					return;
				}
			}
			
			
			for (let i = 0; i < STRATEC_IT_WhiteList.length; i++) {
				if (fromAddress.split('@')[1].trim().toUpperCase() === STRATEC_IT_WhiteList[i].trim().toUpperCase()) {
					legitimateDetected();
					return;
				}
			}
			
			/* for (let i = 0; i < STRATEC_IT_WhiteList_MailAddresses.length; i++) {
				if (fromAddress.trim().toUpperCase() === STRATEC_IT_WhiteList_MailAddresses[i].trim().toUpperCase()) {
					legitimateDetected();
					return;
				}
			} */
			
            // Now that we have a ChangeKey value, we can use EWS to forward the mail item.
            // The first thing we'll do is get an array of email addresses that the user
            // has typed into the To: text box.
            // We'll also get the comment that the user may have provided in the Comment: text box.
            //var toAddresses = document.getElementById("groupEmails").value;
            //var addresses = toAddresses.split(";");
            //var addressesSoap = "";

            // The following loop build an XML fragment that we will insert into the SOAP message
            //for (var address = 0; address < addresses.length; address++) {
            //    addressesSoap += "<t:Mailbox><t:EmailAddress>" + addresses[address] + "</t:EmailAddress></t:Mailbox>";
            //}
			
            var comment = document.getElementById("groupCommentEng").value + document.getElementById("groupCommentGer").value;
			
			var addressesSoap = "<t:Mailbox><t:EmailAddress>" + "security-awareness@stratec.com" + "</t:EmailAddress></t:Mailbox>";
			
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
                '          <t:NewBodyContent BodyType="Text">' + escapeHTML("User comment: " + comment + "\n\nHeaders: " + headerText) + ' </t:NewBodyContent>' +
                '        </t:ForwardItem>' +
                '      </m:Items>' +
                '    </m:CreateItem>' +
                '  </soap:Body>' +
                '</soap:Envelope>';

            // As before, the makeEwsRequestAsync method accepts a string of SOAP and a callback function.
            // The only difference this time is that the body of the SOAP message requests that the item
            // be forwarded (rather than retrieved as in the previous method call)
			
			// app.showNotification("Status", soapToForwardItem);
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
            app.showNotification("Status", asyncResult.error.message);
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
                app.showNotification("Status", "The email was successfully delivered!");
				mailSuccessfullyReported();
            }
            else {
                app.showNotification("Status", "The following error code was recieved: " + result);
            }
        }
    }

	function deleteMail() {
        var item = Office.context.mailbox.item;
        item_id = item.itemId;
        mailbox = Office.context.mailbox;

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
            '        <MoveItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
			'			xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
			'			  <ToFolderId>' +
			'				<t:DistinguishedFolderId Id="deleteditems"/>' +
			'			  </ToFolderId>' +
			'			  <ItemIds>' +
			'				<t:ItemId Id="' + item_id + '"/>' +
			'			  </ItemIds>' +
			'			</MoveItem>' +
            '  </soap:Body>' +
            '</soap:Envelope>';

        // The makeEwsRequestAsync method accepts a string of SOAP and a callback function
		
		// app.showNotification("Status", soapToGetItemData);
        mailbox.makeEwsRequestAsync(soapToGetItemData, soapToDeleteItemCallback);
    }
	
	
	function soapToDeleteItemCallback(asyncResult) {
        var parser;
        var xmlDoc;

        if (asyncResult.error != null) {
            app.showNotification("Status", asyncResult.error.message);
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
                app.showNotification("Status", "The email was successfully moved to the deleted items folder!");
            }
            else {
                app.showNotification("Status", "The following error code was recieved: " + result);
            }
        }
    }
	
})();

// Office.actions.associate("simpleForwardEmail", simpleForwardEmail);
