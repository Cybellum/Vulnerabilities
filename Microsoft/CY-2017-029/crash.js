(function () {
    "use strict";
    // The initialize function is run each time the page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {         
            // Use this to check whether the API is supported in the Word client.
            if (Office.context.requirements.isSetSupported('WordApi', 1.3)) {
                Word.run(function (context) 
                {
                    var body = context.document.body;
                    var myRange = body.getRange();
                    return context.sync().then(function()
                    {
                        myRange.insertFileFromBase64('aaa');
                        return context.sync();
                    })
                    .catch(function (error)
                    {
                    });
                    })
                    .catch(function (error) {
                });
            }
        });
    };	
})();