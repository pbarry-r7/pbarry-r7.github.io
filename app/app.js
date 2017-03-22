var app = (function(){  // jshint ignore:line
  'use strict';

  var self = {};

  // Globals used when processing phish submission to backend
  var reporterMailbox;
  var phishToSubmitItem;
  var phishToSubmitItemIdElement;
  var phishToSubmitMimeContent;
  var subject;
  var reporterEmail;
  var reporterName;
  var retryOnFailureRemaining;
  var apiEndpoint;
  var apiIngestionKey;
  var destEmail;

  // Constants related to messaging
  var RETRY_ON_FAILURE_COUNT = 3;
  var NEW_STATUS_MESSAGE = " This email was submitted as a likely phish to Rapid7";

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
    self.showNotification = function(header, details){
      if (details === undefined) {
        details = "";
      }
      jQuery('#notification-message-header').text(header);
      jQuery('#notification-message-body').text(details);
      jQuery('#notification-message').slideDown('fast');
    };

    self.showErrorNotification = function(text){
      self.showNotification("Oops, an error occurred", text);
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

    function submissionURL(api_endpoint, ingestion_key) {
      return api_endpoint + "/" + ingestion_key;
    }

    // Main entry point for submitting an email as a phish.
    self.submitMessageAsPhish = function(mailbox, userIngestionKey, userApiEndpoint, userDestEmail) {

      app.showNotification("Reporting email as phishing...");

      reporterMailbox = mailbox;
      phishToSubmitItem = reporterMailbox.item;
      apiEndpoint = userApiEndpoint;
      apiIngestionKey = userIngestionKey;
      destEmail = userDestEmail;

      // Values which we want to accompany the actual email content being submitted...
      subject = reporterMailbox.item.subject;
      reporterEmail = reporterMailbox.userProfile.emailAddress;
      reporterName = reporterMailbox.userProfile.displayName;

      // Use GetItem to retreive the message, itself, we're submitting...
      retryOnFailureRemaining = RETRY_ON_FAILURE_COUNT;
      getItem();
      return;
    }

    // Get details on this message, including content.
    function getItem() {
      reporterMailbox.makeEwsRequestAsync(getItemRequestSoap(phishToSubmitItem.itemId), getItemCallback);
    }

    function getItemCallback(result) {
      if (result.error != null) {
        if (result.error.message.indexOf("Response exceeds 1 MB size limit.") == 0) {
          // Message is too big to retreive from Microsoft's add-in API.
          self.showErrorNotification("Please forward this message as an attachment to " + destEmail);
        } else if (retryOnFailureRemaining-- > 0) {
          // Give it another try...
          self.showNotification("Reporting email as phishing..." + Array(RETRY_ON_FAILURE_COUNT - retryOnFailureRemaining + 1).join("."));
          getItem();
          return;
        } else {
          // Out of retries...
          self.showErrorNotification("Please forward this message as an attachment to " + destEmail);
        }
        return;
      }

      var xmlparser = new DOMParser();
      var resultxml = xmlparser.parseFromString(result.value, "text/xml");
      phishToSubmitItemIdElement = resultxml.getElementsByTagName("ItemId")[0];
      if (phishToSubmitItemIdElement) {
        phishToSubmitMimeContent = resultxml.getElementsByTagName("MimeContent")[0].childNodes[0].nodeValue;
      } else {
        // IE and Outlook Desktop require the 't:' prepend...
        phishToSubmitItemIdElement = resultxml.getElementsByTagName("t:ItemId")[0];
        phishToSubmitMimeContent = resultxml.getElementsByTagName("t:MimeContent")[0].childNodes[0].nodeValue;
      }

      retryOnFailureRemaining = RETRY_ON_FAILURE_COUNT;
      postSubmission();
    }

    // Send message to the backend.
    function postSubmission() {
      var POST_TIMEOUT_MS = 5000;  // timeout (ms) between POST retries

      var postData = { "subject": subject, "name": reporterName, "reply_address": reporterEmail,
          "raw": phishToSubmitMimeContent };

      // POST the email to the backend...
      jQuery.ajaxSetup({
        crossDomain: true,
        contentType: "application/json",
        timeout: POST_TIMEOUT_MS
      });
      jQuery.post(submissionURL(apiEndpoint, apiIngestionKey), JSON.stringify(postData), function(result) {
        retryOnFailureRemaining = RETRY_ON_FAILURE_COUNT;
        moveMessage();
      }, 'json')
      .fail(function(result) {
        if (retryOnFailureRemaining-- > 0) {
          self.showNotification("Reporting email as phishing..." + Array(RETRY_ON_FAILURE_COUNT - retryOnFailureRemaining + 1).join("."));
          postSubmission();
          return;
        }
        self.showErrorNotification("Please forward this message as an attachment to " + destEmail);
      }); // API POST
    }

    // Move the message into the Deleted Items folder.
    function moveMessage() {
      reporterMailbox.makeEwsRequestAsync(moveItemRequestSoap(phishToSubmitItemIdElement.getAttribute("Id"),
                                                        phishToSubmitItemIdElement.getAttribute("ChangeKey"),
                                                        "deleteditems"), moveMessageCallback);
    }

    function moveMessageCallback(result) {
      if (result.error != null) {
        if (retryOnFailureRemaining-- > 0) {
          self.showNotification("Reporting email as phishing..." + Array(RETRY_ON_FAILURE_COUNT - retryOnFailureRemaining + 1).join("."));
          moveMessage();
          return;
        }
        retryOnFailureRemaining = RETRY_ON_FAILURE_COUNT;
        replaceMessageStatus();
        self.showNotification("Message reported!", "Please go ahead and delete this message.");
        return;
      }
      retryOnFailureRemaining = RETRY_ON_FAILURE_COUNT;
      replaceMessageStatus();
      self.showNotification("Message reported!", "Please close this window if it does not close automatically.");
    }

    // Add a 'tag' to the message that mentions it was submitted to R7.
    function replaceMessageStatus() {
      phishToSubmitItem.notificationMessages.replaceAsync("status", {
        type: "informationalMessage",
        icon: "icon-16",
        message: NEW_STATUS_MESSAGE,
        persistent: true
      }, {}, replaceMessageStatusCallback);
    }

    function replaceMessageStatusCallback(result) {
      if (result.error != null) {
        if (retryOnFailureRemaining-- > 0) {
          replaceMessageStatus();
        }
        // All out of retries, not a huge deal to give up on tagging it...
      }
      // Status message added (or given up on), now move item to Deleted Items folder...
    }
  };

  return self;
})();
