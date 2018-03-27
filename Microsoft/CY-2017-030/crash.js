(function () {
    "use strict";
    // The initialize function is run each time the page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Use this to check whether the API is supported in the Excel client.
            if (Office.context.requirements.isSetSupported('ExcelApi', 1.6)) {
                Excel.run(function (context) 
                {
                    var range = context.workbook.worksheets.getActiveWorksheet().getRange('A1');
                    return context.sync()
                    .then (function(){
                        range.delete('Up'); 
                        return context.sync();
                        })
                    .catch(function(error){})
                    .then (function(){
                            range.getIntersectionOrNullObject(); 
                            return context.sync();})
                    .catch(function(error){})		
                        })
                    .catch(function (error) {
                });
            }
        });
    };	
})();