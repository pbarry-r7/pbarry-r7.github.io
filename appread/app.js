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


    function replaceMessageStatus(item, message) {
        item.notificationMessages.replaceAsync("status", {
          type: "informationalMessage",
          icon: "icon-16",
          message: message,
          persistent: true
        });
    }

    self.submitMessageAsPhish = function(mailbox, dest_email) {
      var phishToForwardItem = mailbox.item;
      var subject = phishToForwardItem.subject;
      var reporterEmail = mailbox.userProfile.EmailAddress

      // Use GetItem to retreive info on the message we're forwarding...
      mailbox.makeEwsRequestAsync(getItemRequestSoap(phishToForwardItem.itemId), function(result) {
        if (result.error != null) {
          self.showNotification("An error occured", "Please forward this message as an attachment to " + dest_email);
          return;
        }

        var xmlparser = new DOMParser();
        var resultxml = xmlparser.parseFromString(result.value, "text/xml");
        var phishToForwardItemIdElement = resultxml.getElementsByTagName("ItemId")[0];
        var phishToForwardMimeContent;
        if (phishToForwardItemIdElement) {
          phishToForwardMimeContent = resultxml.getElementsByTagName("MimeContent")[0].innerHTML;
        } else {
          // IE and Outlook Desktop require the prepend...
          phishToForwardItemIdElement = resultxml.getElementsByTagName("t:ItemId")[0];
          phishToForwardMimeContent = resultxml.getElementsByTagName("t:MimeContent")[0].innerHTML;
        }
        var subject = mailbox.item.subject;
        var reporterEmail = mailbox.userProfile.emailAddress;
        var reporterName = mailbox.userProfile.displayName;

      }); // Get Item
    }; // submitMessageAsPhish

  };

  return self;
})();
