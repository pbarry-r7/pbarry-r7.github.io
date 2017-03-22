(function(){
  'use strict';

  var params = {};
  var ingestionKey = "";
  var apiEndpoint = "";
  var destEmail = "";

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      var paramsArray = window.location.search.substring(1).split("&");
      while (paramsArray.length) {
        var kvp = paramsArray[0].split("=");
        params[kvp[0]] = kvp[1];
        paramsArray.splice(0,1);
      }

      // If we were passed an ID, use that to load config values...
      if (params.conf != null) {
//        JSON.parse(atob(params.conf), (key, value) => {
//          switch (key) {
//            case 'insight_ingestion_key':
//              ingestionKey = value;
//              break;
//            case 'insight_api_endpoint':
//              apiEndpoint = value;
//              break;
//            case 'insight_dest_email':
//              destEmail = value;
//              break;
//            default:
//              break;
//          }
//        });

        var conf = JSON.parse(Base64.decode(params.conf));
        ingestionKey = conf.insight_ingestion_key;
        apiEndpoint = conf.insight_api_endpoint;
        destEmail = conf.insight_dest_email;
      }

      app.initialize();
      addClickableActions();
    });
  };

  // Add "report" and "info" button/link actions
  function addClickableActions(){
    jQuery('#report_button').click(function() {
      app.submitMessageAsPhish(Office.context.mailbox, ingestionKey, apiEndpoint, destEmail);
    });
    jQuery('#more_info_link').click(function() {
      jQuery('#more_info_text').css('display', 'inline');
    });
  }

})();
