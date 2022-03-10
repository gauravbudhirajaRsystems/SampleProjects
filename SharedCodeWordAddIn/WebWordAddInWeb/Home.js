
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            
            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {

            }

        });
    };
   
})();

function ShowUnicode() {
    Word.run(function (context) {
        const range = context.document.getSelection();
        context.load(range, 'text');
        return context.sync().then(function () {
            const url = "https://localhost:44324/wordanalyzer/unicode?value=" + range.text;
            $.ajax({
                type: "GET",
                url: url,
                success: function (data) {
                    let htmlData = data.replace(/\r\n/g, '<br>');
                    $("#txtUnicodeResult").html(htmlData);
                },
                error: function (data) {
                    $("#txtUnicodeResult").html("error occurred in ajax call.");
                }
            });
        });
    });
}

function ShowCharCount() {
    Word.run(function (context) {
        const range = context.document.getSelection();
        context.load(range, 'text');
        return context.sync().then(function () {
            const url = "https://localhost:44324/wordanalyzer/charcount?value=" + range.text;
            $.ajax({
                type: "GET",
                url: url,
                success: function (data) {
                    let htmlData = data.replace(/\r\n/g, '<br>');
                    $("#txtCharCountResult").html(htmlData);
                },
                error: function (data) {
                    $("#txtCharCountResult").html("error occurred in ajax call.");
                }
            });
        });
    });
}

function ShowWordCount() {
    Word.run(function (context) {
        const range = context.document.getSelection();
        context.load(range, 'text');
        return context.sync().then(function () {
            const url = "https://localhost:44324/wordanalyzer/wordcount?value=" + range.text;
            $.ajax({
                type: "GET",
                url: url,
                success: function (data) {
                    let htmlData = data.replace(/\r\n/g, '<br>');
                    $("#txtWordCountResult").html(htmlData);
                },
                error: function (data) {
                    $("#txtWordCountResult").html("error occurred in ajax call.");
                }
            });
        });
    });
}