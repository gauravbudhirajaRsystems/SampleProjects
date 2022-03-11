/* eslint-disable office-addins/load-object-before-read */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-empty */
/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import { base64Image } from "../../base64Image";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    // eslint-disable-next-line no-empty
    if (Office.context.requirements.isSetSupported("WordApi", "1.1")) {
      // for Office 2016
      document.getElementById("show-unicode").onclick = ShowUnicode;
      document.getElementById("show-charcount").onclick = ShowCharCount;
      document.getElementById("show-wordcount").onclick = ShowWordCount;
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
    } else if (Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      // For Office 2019
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
    } else {
    }

    // Assign event handlers and other initialization logic.
    // document.getElementById("insert-paragraph").onclick = insertParagraph;
    // //document.getElementById("apply-style").onclick = insertParagraphNew;
    // document.getElementById("apply-style").onclick = applyStyle;
    // document.getElementById("apply-custom-style").onclick = applyCustomStyle;
    // document.getElementById("change-font").onclick = changeFont;
    // document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;
    // document.getElementById("insert-text-outside-range").onclick = insertTextBeforeRange;
    // document.getElementById("replace-text").onclick = replaceText;
    // document.getElementById("insert-image").onclick = insertImage;
    // document.getElementById("insert-html").onclick = insertHTML;
    // document.getElementById("insert-table").onclick = insertTable;
    // document.getElementById("create-content-control").onclick = createContentControl;
    // document.getElementById("replace-content-in-control").onclick = replaceContentInControl;

    // document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    //document.getElementById("run").onclick = run;
  }
});

// async function replaceContentInControl() {
//   await Word.run(async (context) => {
//     const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
//     serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");

//     await context.sync();
//   }).catch(function (error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// async function createContentControl() {
//   await Word.run(async (context) => {
//     const serviceNameRange = context.document.getSelection();
//     const serviceNameContentControl = serviceNameRange.insertContentControl();
//     serviceNameContentControl.title = "Service Name";
//     serviceNameContentControl.tag = "serviceName";
//     serviceNameContentControl.appearance = "Tags";
//     serviceNameContentControl.color = "blue";

//     await context.sync();
//   }).catch(function (error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// async function insertTable() {
//   await Word.run(async (context) => {
//     //const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
//     const secondParagraph = context.document.body.paragraphs.getFirst();
//     const tableData = [
//       ["Name", "ID", "Birth City"],
//       ["Bob", "434", "Chicago"],
//       ["Sue", "719", "Havana"],
//     ];
//     secondParagraph.insertTable(3, 3, "After", tableData);
//     //context.document.body.paragraphs.getFirst().getNext().insertTable(3, 3, "After", tableData);
//     //context.document.body.insertTable(3, 3, "Start", tableData);

//     await context.sync();
//   }).catch(function (error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// async function insertHTML() {
//   await Word.run(async (context) => {
//     //const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
//     context.document.body.insertHtml(
//       "<p style='font-family: verdana;'>Inserted HTML.</p><p>Another paragraph</p>",
//       "Start"
//     );

//     await context.sync();
//   }).catch(function (error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// async function insertImage() {
//   await Word.run(async (context) => {
//     context.document.body.insertInlinePictureFromBase64(base64Image, "End");

//     await context.sync();
//   }).catch(function (error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// async function replaceText() {
//   await Word.run(async (context) => {
//     const doc = context.document;
//     const originalRange = doc.getSelection();
//     originalRange.insertText("many", "Replace");

//     await context.sync();
//   }).catch(function (error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// async function insertTextBeforeRange() {
//   await Word.run(async (context) => {
//     const doc = context.document;
//     const originalRange = doc.getSelection();
//     originalRange.insertText("Office 2019, ", "Before");

//     originalRange.load("text");
//     await context.sync();

//     doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");

//     await context.sync();
//   }).catch(function (error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// // export async function run() {
// //   return Word.run(async (context) => {
// //     /**
// //      * Insert your Word code here
// //      */

// //     // insert a paragraph at the end of the document.
// //     const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

// //     // change the paragraph color to blue.
// //     paragraph.font.color = "blue";

// //     await context.sync();
// //   });
// // }

// // eslint-disable-next-line @typescript-eslint/no-unused-vars

// async function applyStyle() {
//   await Word.run(async (context) => {
//     const firstParagraph = context.document.body.paragraphs.getFirst();
//     firstParagraph.styleBuiltIn = Word.Style.intenseQuote;

//     await context.sync();
//   }).catch(function (error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// async function applyCustomStyle() {
//   await Word.run(async (context) => {
//     const lastParagraph = context.document.body.paragraphs.getLast();
//     lastParagraph.style = "MyCustomStyle";
//     await context.sync();
//   }).catch(function (error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// async function changeFont() {
//   await Word.run(async (context) => {
//     const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
//     secondParagraph.font.set({
//       name: "Courier New",
//       bold: true,
//       size: 18,
//     });

//     await context.sync();
//   }).catch(function (error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// async function insertTextIntoRange() {
//   await Word.run(async (context) => {
//     const doc = context.document;
//     const originalRange = doc.getSelection();
//     originalRange.insertText(" (C2R)", "End");

//     originalRange.load("text");
//     await context.sync();

//     doc.body.insertParagraph("Original range: " + originalRange.text, "End");

//     await context.sync();
//   }).catch(function (error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// async function insertParagraph() {
//   await Word.run(async (context) => {
//     const docBody = context.document.body;
//     docBody.insertParagraph(
//       "Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
//       "Start"
//     );
//     await context.sync();
//   }).catch(function (error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// APi Interaction Code Starts

function ShowUnicode() {
  Word.run(function (context) {
    var originalRange = context.document.getSelection();
    context.load(originalRange, "text");
    return context.sync().then(function () {
      var url = "https://localhost:44325/wordanalyzer/unicode?value=" + originalRange.text;
      $.ajax({
        type: "GET",
        url: url,
        success: function (data_1) {
          let htmlData = data_1.replace(/\r\n/g, "<br>");
          $("#txtUnicodeResult").html(htmlData);
        },
        error: function (data_3) {
          $("#txtUnicodeResult").html("error occurred in ajax call.");
        },
      });
    });
  }).catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function ShowCharCount() {
  Word.run(function (context) {
    const originalRange = context.document.getSelection();
    context.load(originalRange, "text");
    return context.sync().then(function () {
      const url = "https://localhost:44325/wordanalyzer/charcount?value=" + originalRange.text;
      $.ajax({
        type: "GET",
        url: url,
        success: function (data_1) {
          let htmlData = data_1.replace(/\r\n/g, "<br>");
          $("#txtCharCountResult").html(htmlData);
        },
        error: function (data_3) {
          $("#txtCharCountResult").html("error occurred in ajax call.");
        },
      });
    });
  });
}

function ShowWordCount() {
  Word.run(function (context) {
    const originalRange = context.document.getSelection();
    context.load(originalRange, "text");
    return context.sync().then(function () {
      const url = "https://localhost:44325/wordanalyzer/wordcount?value=" + originalRange.text;
      $.ajax({
        type: "GET",
        url: url,
        success: function (data_1) {
          let htmlData = data_1.replace(/\r\n/g, "<br>");
          $("#txtWordCountResult").html(htmlData);
        },
        error: function (data_3) {
          $("#txtWordCountResult").html("error occurred in ajax call.");
        },
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
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function alignTextRight() {
  Word.run(function (context) {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    var firstParagraph;
    return context
      .sync()
      .then(function () {
        if (paragraphs.items.length > 0) {
          firstParagraph = paragraphs.items[0];
          firstParagraph.load("alignment");
        }
      })
      .then(context.sync)
      .then(function () {
        firstParagraph.alignment = Word.Alignment.right;
      });
  }).catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function applyInBuiltStyle() {
  Word.run(function (context) {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    var firstParagraph;
    return context
      .sync()
      .then(function () {
        if (paragraphs.items.length > 0) {
          firstParagraph = paragraphs.items[0];
          firstParagraph.load("styles");
        }
      })
      .then(context.sync)
      .then(function () {
        firstParagraph.styleBuiltIn = "Emphasis";
      });
  }).catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function changeFont() {
  Word.run(function (context) {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    var secondParagraph;
    return context
      .sync()
      .then(function () {
        if (paragraphs.items.length > 1) {
          secondParagraph = paragraphs.items[1];
          secondParagraph.load("font");
        }
      })
      .then(context.sync)
      .then(function () {
        var value = secondParagraph;
        secondParagraph.font.set({
          name: "Courier New",
          bold: true,
          size: 18,
        });
      });
  }).catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function insertTextIntoRange() {
  Word.run(function (context) {
    var originalRange = context.document.getSelection();
    context.load(originalRange, "text");
    return context
      .sync()
      .then(function () {
        originalRange.insertText(" (C2R)", "End");
        originalRange.load("text");
      })
      .then(context.sync)
      .then(function () {
        context.document.body.insertParagraph("Original range: " + originalRange.text, "End");
      });
  }).catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function insertTextBeforeRange() {
  Word.run(function (context) {
    var originalRange = context.document.getSelection();
    context.load(originalRange, "text");
    return context
      .sync()
      .then(function () {
        originalRange.insertText("Office 2016, ", "Before");
        originalRange.load("text");
      })
      .then(context.sync)
      .then(function () {
        context.document.body.insertParagraph("Original range: " + originalRange.text, "End");
      });
  }).catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function replaceText() {
  Word.run(function (context) {
    var originalRange = context.document.getSelection();
    context.load(originalRange, "text");
    return context.sync().then(function () {
      originalRange.insertText("many", "Replace");
    });
  }).catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function insertImage() {
  Word.run(function (context) {
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    return context.sync();
  }).catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
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
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
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
    });
  }).catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function createContentControl() {
  Word.run(function (context) {
    let serviceNameRange = context.document.getSelection();
    context.load(serviceNameRange, "text");
    return context.sync().then(function () {
      let serviceNameContentControl = serviceNameRange.insertContentControl();
      serviceNameContentControl.title = "Service Name";
      serviceNameContentControl.tag = "serviceName";
      serviceNameContentControl.appearance = "Tags";
      serviceNameContentControl.color = "blue";
    });
  }).catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
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
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

// Document Interaction Code Ends
