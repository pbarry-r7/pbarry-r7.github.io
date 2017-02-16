(function(){
  'use strict';

  var params = {}
  var ingestion_key = ""
  var dest_email = ""

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      var params_array = window.location.search.substring(1).split("&");
      while (params_array.length) {
        var kvp = params_array[0].split("=");
        params[kvp[0]] = kvp[1];
        params_array.splice(0,1);
      }
      if (params.id != null) {
        jQuery.getScript("../../conf/"+params.id+".js", function(){
          ingestion_key = as_ingestion_key;
          dest_email = as_dest_email;
        });
      }

      app.initialize();

      displayItemDetails();
      addClickableActions();
    });
  };

  // Displays the "Subject" and "From" fields, based on the current mail item
  function displayItemDetails(){
    var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
    jQuery('#subject').text(item.subject);

    var from;
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      from = Office.cast.item.toMessageRead(item).from;
    } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      from = Office.cast.item.toAppointmentRead(item).organizer;
    }

    if (from) {
      jQuery('#from').text(from.displayName);
      jQuery('#from').click(function(){
        app.showNotification(from.displayName, from.emailAddress);
      });
    }
  }

  // Add "report" and "cancel" button actions
  function addClickableActions(){
    jQuery('#report_button').click(function() {
      app.showNotification("Reporting email as phishing...");

      app.submitMessageAsPhish(Office.context.mailbox, ingestion_key, dest_email);
    });
    jQuery('#more_info_link').click(function() {
      jQuery('#more_info_text').css('display', 'inline');
    });
  }

})();
