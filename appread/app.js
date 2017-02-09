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

    function createMessageRequestSoap(to, subject, message) {
      message = message || '';
      return '<?xml version="1.0" encoding="utf-8"?>' +
             '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
             '  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
             '  <soap:Body>' +
             '    <CreateItem MessageDisposition="SaveOnly" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
             '      <SavedItemFolderId>' +
             '        <t:DistinguishedFolderId Id="drafts" />' +
             '      </SavedItemFolderId>' +
             '      <Items>' +
             '        <t:Message>' +
             '          <t:ItemClass>IPM.Note</t:ItemClass>' +
             '          <t:Subject>' + subject + '</t:Subject>' +
             '          <t:Body BodyType="Text">' + message + '</t:Body>' +
             '          <t:ToRecipients>' +
             '            <t:Mailbox>' +
             '              <t:EmailAddress>' + to + '</t:EmailAddress>' +
             '            </t:Mailbox>' +
             '          </t:ToRecipients>' +
             '          <t:IsRead>false</t:IsRead>' +
             '        </t:Message>' +
             '      </Items>' +
             '    </CreateItem>' +
             '  </soap:Body>' +
             '</soap:Envelope>';
    } 

    function createAttachmentRequestSoap(itemId, changeKey, name, mimeContent) {
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
             '    <CreateAttachment' +
             '      xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
             '      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
             '      <ParentItemId Id="' + itemId + '" ChangeKey="' + changeKey + '" />' +
             '      <Attachments>' +
             '        <t:ItemAttachment>' +
             '          <t:Name>' + name + '</t:Name>' +
             '          <t:IsInline>false</t:IsInline>' +
             '          <t:Message>' +
             '            <t:MimeContent CharacterSet="UTF-8">' + mimeContent + '</t:MimeContent>' +
             '          </t:Message>' +
             '        </t:ItemAttachment>' +
             '      </Attachments>' +
             '    </CreateAttachment>' +
             '  </soap:Body>' +
             '</soap:Envelope>';
    }

    function sendItemRequestSoap(itemId, changeKey) {
      return '<?xml version="1.0" encoding="utf-8"?>' +
             '<soap:Envelope' +
             '  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
             '  xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
             '  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
             '  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
             '  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
             '  <soap:Body>' +
             '    <m:SendItem SaveItemToFolder="true">' +
             '      <m:ItemIds>' +
             '        <t:ItemId Id="' + itemId + '" ChangeKey="' + changeKey + '"/>' +
             '      </m:ItemIds>' +
             '      <m:SavedItemToFolderId>' +
             '        <t:DistinguishedFolderId Id="sentitems" />' +
             '      </m:SavedItemToFolderId>' +
             '    </m:SendItem>' +
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

    self.submitMessageAsPhish = function(mailbox, dest_email) {
      var phishToForwardItem = mailbox.item;

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

        // Create new message
        mailbox.makeEwsRequestAsync(createMessageRequestSoap(dest_email, "FW: " + phishToForwardItem.subject), function(result) {

          if (result.error != null) {
            self.showNotification("An error occured", "Please forward this message as an attachment to " + dest_email);
            return;
          }

          var xmlparser = new DOMParser();
          var resultxml = xmlparser.parseFromString(result.value, "text/xml");
          var newMessageItemIdElement = resultxml.getElementsByTagName("ItemId")[0];
          if (newMessageItemIdElement == null) {
            // IE and Outlook Desktop require the prepend...
            newMessageItemIdElement = resultxml.getElementsByTagName("t:ItemId")[0];
          }

          // TODO THIS IS WHERE WE'D DO THE ATTACHMENT, IF ONLY MSFT WOULD ALLOW US TO DO SO...  :(

          // Send email off for processing
          mailbox.makeEwsRequestAsync(sendItemRequestSoap(newMessageItemIdElement.getAttribute("Id"),
                                                          newMessageItemIdElement.getAttribute("ChangeKey")),
                                                          function(result) {
            if (result.error != null) {
              self.showNotification("An error occured", "Please forward this message as an attachment to " + dest_email);
              return;
            }

            replaceMessageStatus(phishToForwardItem, " This email was submitted as a likely phish to Rapid7");

            // Delete existing message
            mailbox.makeEwsRequestAsync(moveItemRequestSoap(phishToForwardItemIdElement.getAttribute("Id"),
                                                            phishToForwardItemIdElement.getAttribute("ChangeKey"),
                                                            "deleteditems"), function(result) {
              if (result.error != null) {
                self.showNotification("An error occured", "Please go ahead and delete this message.");
              }
            }); // Move Item
          }); // Send Item
        }); // Create Item
      }); // Get Item
    }; // submitMessageAsPhish

  };

  return self;
})();
