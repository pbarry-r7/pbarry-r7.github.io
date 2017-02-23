var app = (function(){  // jshint ignore:line
  'use strict';

  var self = {};

  // Common initialization function (to be called from each page)
  self.initialize = function(){
    jQuery('body').append(
      '<div id="notification-message">' +
      '<div class="padding">' +
      '<div id="notification-message-close"></div>' +
      '<div id="notification-message-header"></div>' +
      '<div id="notification-message-body"></div>' +
      '</div>' +
      '</div>');

    jQuery('#notification-message-close').click(function(){
      jQuery('#notification-message').hide();
    });

    // After initialization, expose common notification functions
    self.showNotification = function(header, text){
      jQuery('#notification-message-header').text(header);
      jQuery('#notification-message-body').text(text);
      jQuery('#notification-message').slideDown('fast');
    };

    self.closeNotification = function(){
      jQuery('#notification-message').hide();
    };

    function getItemRequestSoap(itemId) {
      return '<?xml version="1.0" encoding="utf-8"?>' +
             '<soap:Envelope' +
             '  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
             '  xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
             '  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
             '  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
             '  <soap:Header>' +
             '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
             '  </soap:Header>' +
             '  <soap:Body>' +
             '    <GetItem' +
             '      xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
             '      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
             '      <ItemShape>' +
             '        <t:BaseShape>Default</t:BaseShape>' +
             '        <t:IncludeMimeContent>true</t:IncludeMimeContent>' +
             '      </ItemShape>' +
             '      <ItemIds>' +
             '        <t:ItemId Id="' + itemId + '" />' +
             '      </ItemIds>' +
             '    </GetItem>' +
             '  </soap:Body>' +
             '</soap:Envelope>';
    }

    function moveItemRequestSoap(itemId, changeKey, destFolder) {
      return '<?xml version="1.0" encoding="utf-8"?>' +
             '<soap:Envelope' +
             '  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
             '  xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
             '  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
             '  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
             '  <soap:Header>' +
             '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
             '  </soap:Header>' +
             '  <soap:Body>' +
             '    <MoveItem' +
             '      xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
             '      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
             '      <ToFolderId>' +
             '        <t:DistinguishedFolderId Id="' + destFolder + '"/>' +
             '      </ToFolderId>' +
             '      <ItemIds>' +
             '        <t:ItemId Id="' + itemId + '" ChangeKey="' + changeKey + '"/>' +
             '      </ItemIds>' +
             '    </MoveItem>' +
             '  </soap:Body>' +
             '</soap:Envelope>';
    }

    function replaceMessageStatus(item, message) {
      item.notificationMessages.replaceAsync("status", {
        type: "informationalMessage",
        icon: "icon-16",
        message: message,
        persistent: true
      });
    }

    function submissionURL(ingestion_key) {
      return "http://localhost:8084/mail/" + ingestion_key;
      //return "http://127.0.0.1:8084/mail/" + ingestion_key;
    }

    self.submitMessageAsPhish = function(mailbox, ingestion_key, dest_email) {
      var phishToSubmitItem = mailbox.item;
      var subject = phishToSubmitItem.subject;
      var reporterEmail = mailbox.userProfile.EmailAddress

      // Use GetItem to retreive info on the message we're submitting...
      mailbox.makeEwsRequestAsync(getItemRequestSoap(phishToSubmitItem.itemId), function(result) {
        if (result.error != null) {
          self.showNotification("An error occured", "Please forward this message as an attachment to " + dest_email);
          return;
        }

        var xmlparser = new DOMParser();
        var resultxml = xmlparser.parseFromString(result.value, "text/xml");
        var phishToSubmitItemIdElement = resultxml.getElementsByTagName("ItemId")[0];
        var phishToSubmitMimeContent;
        if (phishToSubmitItemIdElement) {
          phishToSubmitMimeContent = resultxml.getElementsByTagName("MimeContent")[0].innerHTML;
        } else {
          // IE and Outlook Desktop require the prepend...
          phishToSubmitItemIdElement = resultxml.getElementsByTagName("t:ItemId")[0];
          phishToSubmitMimeContent = resultxml.getElementsByTagName("t:MimeContent")[0].innerHTML;
        }
        var subject = mailbox.item.subject;
        var reporterEmail = mailbox.userProfile.emailAddress;
        var reporterName = mailbox.userProfile.displayName;

        // TODO submit the message to the new endpoint here...!
        // submissionUrl, 
        //self.showNotification("JSON is { \"subject\": \"" + subject + "\", \"name\": \"" + reporterName + "\", \"reply_address\": \"" + reporterEmail + "\", \"raw\": \"" + phishToSubmitMimeContent + "\"}");
        var postData = { "subject": subject, "name": reporterName, "reply_address": reporterEmail, "raw": phishToSubmitMimeContent};

        jQuery.post(submissionURL(ingestion_key), postData, function(result) {

          replaceMessageStatus(phishToSubmitItem, " This email was submitted as a likely phish to Rapid7");

          // Delete existing message
//          mailbox.makeEwsRequestAsync(moveItemRequestSoap(phishToSubmitItemIdElement.getAttribute("Id"),
//                                                          phishToSubmitItemIdElement.getAttribute("ChangeKey"),
//                                                          "deleteditems"), function(result) {
//            if (result.error != null) {
//              self.showNotification("An error occured", "Please go ahead and delete this message.");
//            }

            //self.showNotification("Message reported!", "Please close this window if it does not close automatically.");
//          }); // Move Item
        });  // API POST
      }); // Get Item
    }; // submitMessageAsPhish
  };

  return self;
})();
