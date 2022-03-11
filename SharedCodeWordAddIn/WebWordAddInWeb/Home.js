
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            
            // If not using Word 2016, use fallback logic.
            if (Office.context.requirements.isSetSupported('WordApi', '1.1')) {

                document.getElementById("insert-paragraph").onclick = insertParagraph;
                document.getElementById("align-text-right").onclick = alignTextRight;
                document.getElementById("apply-inbuild-style").onclick = applyInBuiltStyle;
                document.getElementById("change-font").onclick = changeFont;
                document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;
                document.getElementById("insert-text-outside-range").onclick = insertTextBeforeRange;
                document.getElementById("replace-text").onclick = replaceText;
                document.getElementById("insert-image").onclick = insertImage;
                document.getElementById("insert-html").onclick = insertHTML;
                document.getElementById("insert-table").onclick = insertTable;
                document.getElementById("create-content-control").onclick = createContentControl;
                document.getElementById("replace-content-in-control").onclick = replaceContentInControl;
                document.getElementById("supportedVersion").innerHTML = "This code is using Word 2016 or later.";
            }
            else {
                document.getElementById("supportedVersion").innerHTML = "This code is using Word 2016 or later.";

            }
        });
    };
   
})();

// APi Interaction Code Starts


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


// APi Interaction Code Ends



// Document Interaction Code Starts


function insertParagraph() {

    Word.run(function (context) {

        const docBody = context.document.body;
        docBody.insertParagraph(
            "Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
            "Start"
        );

        return context.sync();

    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function alignTextRight() {

    Word.run(function (context) {
        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        var firstParagraph;
        return context.sync().then(function () {
            if (paragraphs.items.length > 0) {
                firstParagraph = paragraphs.items[0];
                firstParagraph.load("alignment");
            }
        }).then(context.sync).then(function () {
            firstParagraph.alignment = Word.Alignment.right;
        })
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function applyInBuiltStyle() {

    Word.run(function (context) {
        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        var firstParagraph;
        return context.sync().then(function () {
            if (paragraphs.items.length > 0) {
                firstParagraph = paragraphs.items[0];
                firstParagraph.load("styles");
            }
        }).then(context.sync).then(function () {
            firstParagraph.styleBuiltIn = "Emphasis";
        })
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function changeFont() {

    Word.run(function (context) {
        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        var secondParagraph;
        return context.sync().then(function () {
            if (paragraphs.items.length > 1) {
                secondParagraph = paragraphs.items[1];
                secondParagraph.load("font");
            }
        }).then(context.sync).then(function () {
            var value = secondParagraph;
            secondParagraph.font.set({
                name: "Courier New",
                bold: true,
                size: 18,
            });
        })
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function insertTextIntoRange() {

    Word.run(function (context) {
        var originalRange = context.document.getSelection();
        context.load(originalRange, 'text');
        return context.sync().then(function () {
            originalRange.insertText(" (C2R)", "End");
            originalRange.load("text");

        }).then(context.sync).then(function () {
            context.document.body.insertParagraph("Original range: " + originalRange.text, "End");
        })
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function insertTextBeforeRange() {

    Word.run(function (context) {
        var originalRange = context.document.getSelection();
        context.load(originalRange, 'text');
        return context.sync().then(function () {
            originalRange.insertText("Office 2016, ", "Before");
            originalRange.load("text");

        }).then(context.sync).then(function () {
            context.document.body.insertParagraph("Original range: " + originalRange.text, "End");
        })
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function replaceText() {

    Word.run(function (context) {
        var originalRange = context.document.getSelection();
        context.load(originalRange, 'text');
        return context.sync().then(function () {
            originalRange.insertText("many", "Replace");
        });
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function insertImage() {

    Word.run(function (context) {
        context.document.body.insertInlinePictureFromBase64(base64Image, "End");
        return context.sync();
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function insertHTML() {

    Word.run(function (context) {
        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        var blankParagraph;
        return context.sync().then(function () {
            if (paragraphs.items.length > 0) {
                blankParagraph = paragraphs.items[paragraphs.items.length - 1].insertParagraph("", "After");
                blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
            }
        });
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function insertTable() {

    Word.run(function (context) {
        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        var blankParagraph;
        const tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
        return context.sync().then(function () {
            if (paragraphs.items.length > 0) {
                //blankParagraph = paragraphs.items[paragraphs.items.length - 1].insertParagraph("", "After");
                blankParagraph = paragraphs.items[paragraphs.items.length - 1];
                blankParagraph.insertTable(3, 3, "After", tableData);
            }
        })
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function createContentControl() {

    Word.run(function (context) {
        let serviceNameRange = context.document.getSelection();
        context.load(serviceNameRange, 'text');
        return context.sync().then(function () {
            let serviceNameContentControl = serviceNameRange.insertContentControl();
            serviceNameContentControl.title = "Service Name";
            serviceNameContentControl.tag = "serviceName";
            serviceNameContentControl.appearance = "Tags";
            serviceNameContentControl.color = "blue";
        });
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function replaceContentInControl() {

    Word.run(function (context) {
        let doc = context.document;
        doc.load("contentControls");
        return context.sync().then(function () {
            var serviceNameContentControl = doc.contentControls.getByTag("serviceName").items[0];
            serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
        });
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}


// Document Interaction Code Ends