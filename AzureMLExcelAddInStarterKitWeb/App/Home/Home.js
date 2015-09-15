/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);
            $('#runPrediction').click(runPrediction);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }
    function JSONize(str) {
        return str
          // wrap keys without quote with valid double quote
          .replace(/([\$\w]+)\s*:/g, function (_, $1) { return '"' + $1 + '":' })
          // replacing single quote wrapped ones to double quote 
          .replace(/'([^']+)'/g, function (_, $1) { return '"' + $1 + '"' })
    }

    function runPrediction() {
        // create a new Request Context
        var ctx = new Excel.RequestContext();

        jQuery.support.cors = true;
        
        var inputData = '{"Inputs": {"input1": {ColumnNames: ["sentiment_label","tweet_text"],"Values": [ ["0", "0" ], ["0", "good"] ] }}, "GlobalParameters": {}}';
        var ajaxData = JSON.stringify(inputData);

        var serviceUrl = "https://ussouthcentral.services.azureml.net/workspaces/b21d05124424412490e9ae520d287651/services/68b488f035d54fc2918689b2f750a3c1/score";
        var serviceUrl = "https://ussouthcentral.services.azureml.net/workspaces/b21d05124424412490e9ae520d287651/services/68b488f035d54fc2918689b2f750a3c1/execute?api-version=2.0&details=true";
        $.ajax({
            type: "POST",
            url: serviceUrl,
            data: ajaxData,
            headers: {
                "Authorization": "Bearer 3qZX7fJovhn6Nada84+Vi/WVSaGYFTXBVRYziJsb/Pm6OCZJ0iehYnyYLonKVluLnUwje96Q+nKHaxiGqdycIw==",
                "Content-Type": "application/json;charset=utf-8"
            }
        }).done(function (data) {
            console.log(data);
        });

    }

    // Helper method
    
})();